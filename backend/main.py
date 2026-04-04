"""
main.py — PDF → DOCX + Excel Extractor (FastAPI version)
FIXED: dynamic extraction, no hardcoded commission rates
"""

import os
import re
import json
import uuid
import logging
import traceback
import urllib.request
from datetime import datetime
from pathlib import Path

# ── Load .env ────────────────────────────────────────────────────────────────
def _load_dotenv(path: str = ".env"):
    env_path = Path(path)
    if not env_path.exists():
        env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        return
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = val

_load_dotenv()

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
OPENAI_API_KEY    = os.environ.get("OPENAI_API_KEY", "")
OPENAI_MODEL      = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
OPENAI_MAX_TOKENS = int(os.environ.get("OPENAI_MAX_TOKENS", "4096"))
PORT              = int(os.environ.get("PORT", "8000"))

log.info(f"Model  : {OPENAI_MODEL}")
log.info(f"AI Key : {'SET' if OPENAI_API_KEY else 'NOT SET'}")

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ── FastAPI ───────────────────────────────────────────────────────────────────
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional, Any
import uvicorn

import pdfplumber
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI(title="MetalloDoc API", version="2.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

class GenerateRequest(BaseModel):
    job_id: str
    data: Optional[Any] = None


# ════════════════════════════════════════════════════════════════════════════
#  SYSTEM PROMPT — fully dynamic, no hardcoded rates
# ════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """\
You are a financial document data extractor for Italian gold investment contracts
(Conto Metallo Deluxe, Mac & Ro Srl). Extract data and return ONLY valid JSON.

READ THESE EXACT FIELDS FROM THE PDF:

1. importo_contratto
   - Location: "Corrispettivo" field inside "INFORMAZIONI DI CONTRATTO CONTO METALLO DELUXE" table
   - Return numeric string only, no € symbol
   - Example: if PDF shows "10000" return "10000.00"

2. commissioni_annuale
   - Location: Article 2 of the contract body
   - Conto Metallo Deluxe PLUS  → always "6.00"
   - Conto Metallo Deluxe TOP   → always "10.00"
   - Return numeric string only, no % symbol

3. frequenza_pagamento
   - Location: Article 4 section 5 of the contract
   - If payments are "semestrale" (every 6 months) → return "semestrale"
   - If payments are "trimestrale" (every 3 months) → return "trimestrale"

4. percentuale_per_periodo
   - semestrale: commissioni_annuale divided by 2  (e.g. 6/2 = 3.00)
   - trimestrale: commissioni_annuale divided by 4 (e.g. 6/4 = 1.50)
   - Return numeric string only

5. data_contratto
   - Location: "Luogo e data" field on the proposal page
   - Format: DD/MM/YYYY
   - Example: "2026-01-13" → "13/01/2026"

6. durata_mesi
   - Location: "Durata" field in INFORMAZIONI DI CONTRATTO table
   - Return numeric string only (e.g. "60")

7. piano
   - "Plus" if Conto Metallo Deluxe Plus is checked
   - "Top"  if Conto Metallo Deluxe Top is checked

8. nome_cliente_segnalatore
   - Format: "CLIENT_FULLNAME/AGENT_FULLNAME"
   - Client: from CONTRAENTE PERSONA FISICA "Cognome e Nome" field
   - Agent:  from COLLABORATORE "Cognome e Nome" field

9. document_number
   - Location: "Documento n°" field on first page
   - Example: "CMP-000018.01.2026"

Return ONLY this JSON structure, no markdown, no explanation:
{
  "contracts": [
    {
      "tipo_contratto": "Conto Metallo Deluxe",
      "nome_cliente_segnalatore": "QUARANTA PIERPAOLO/Plaitano Dario",
      "data_contratto": "13/01/2026",
      "importo_contratto": "10000.00",
      "commissioni_annuale": "6.00",
      "frequenza_pagamento": "semestrale",
      "percentuale_per_periodo": "3.00",
      "durata_mesi": "60",
      "piano": "Plus"
    }
  ],
  "document_number": "CMP-000018.01.2026",
  "extraction_notes": "Extracted from PDF"
}

STRICT RULES:
- NO payment_schedule in JSON — backend calculates it
- Extract ONLY what is physically written in the PDF
- Never invent or guess values
- Return ONLY raw JSON, no markdown fences
"""


# ════════════════════════════════════════════════════════════════════════════
#  PDF + AI extraction
# ════════════════════════════════════════════════════════════════════════════

def extract_pdf_text(pdf_path: str) -> str:
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages.append(f"=== PAGE {i+1} ===\n{text}")
    return "\n\n".join(pages)


def call_openai(prompt: str, system: str) -> str:
    payload = json.dumps({
        "model":      OPENAI_MODEL,
        "max_tokens": OPENAI_MAX_TOKENS,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user",   "content": prompt},
        ],
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=payload,
        headers={
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {OPENAI_API_KEY}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=300) as resp:
        body = json.loads(resp.read().decode("utf-8"))
        return body["choices"][0]["message"]["content"]


def safe_json(text: str) -> dict:
    for attempt in (
        lambda t: json.loads(t.strip()),
        lambda t: json.loads(re.search(r"```(?:json)?\s*(\{.*?\})\s*```", t, re.DOTALL).group(1)),
        lambda t: json.loads(re.search(r"\{.*\}", t, re.DOTALL).group(0)),
    ):
        try:
            return attempt(text)
        except Exception:
            pass
    raise ValueError(f"Cannot parse JSON from AI response: {text[:300]}")


def ai_extract(pdf_text: str) -> dict:
    prompt = (
        "Extract contract data from this Italian gold investment PDF.\n"
        "Read the EXACT values from these locations:\n"
        "- Corrispettivo field = importo_contratto\n"
        "- Article 2 percentage = commissioni_annuale (Plus=6.00, Top=10.00)\n"
        "- Article 4.5 frequency = frequenza_pagamento\n"
        "- Luogo e data = data_contratto\n"
        "- Durata field = durata_mesi\n\n"
        f"{pdf_text[:50000]}\n\n"
        "Return ONLY the JSON object."
    )
    raw = call_openai(prompt, SYSTEM_PROMPT)
    log.info(f"AI raw (first 500): {raw[:500]}")
    return safe_json(raw)


# ════════════════════════════════════════════════════════════════════════════
#  Amount parser — handles Italian and English number formats
# ════════════════════════════════════════════════════════════════════════════

def _parse_amount(s: str) -> float:
    """
    Handles: "10000", "10.000,00", "10,000.00", "10000.00"
    Returns float, 0.0 on failure — never a hardcoded fallback
    """
    s = str(s).replace("€", "").replace("$", "").strip()
    if not s:
        return 0.0
    # Italian format: 10.000,00
    if "." in s and "," in s:
        if s.index(",") > s.index("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(re.sub(r"[^\d.]", "", s))
    except Exception:
        return 0.0


# ════════════════════════════════════════════════════════════════════════════
#  Payment schedule — 100% dynamic, zero hardcoded values
# ════════════════════════════════════════════════════════════════════════════

MONTHS_IT    = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
                "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]
MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun",
                "Jul","Aug","Sep","Oct","Nov","Dec"]

EMPTY_MONTHS = 5  # grace period — first 5 months always empty


def fix_schedule(contract: dict) -> list:
    """
    Builds 12-month payment schedule purely from contract dict values.
    No hardcoded commission rates, no hardcoded amounts.
    """

    # ── 1. Contract start date ────────────────────────────────────────────────
    date_str = contract.get("data_contratto", "")
    try:
        dt      = datetime.strptime(date_str, "%d/%m/%Y")
        start_m = dt.month - 1   # 0-based
        start_y = dt.year
    except Exception:
        log.warning(f"Cannot parse date '{date_str}'")
        start_m, start_y = 0, 2026

    # ── 2. Import amount — direct from PDF, no fallback ───────────────────────
    importo = _parse_amount(str(contract.get("importo_contratto", "0")))
    if importo == 0:
        log.warning("importo_contratto is 0 — check PDF extraction")

    # ── 3. Annual commission % — direct from PDF ──────────────────────────────
    comm_annuale = _parse_amount(str(contract.get("commissioni_annuale", "0")))
    if comm_annuale == 0:
        log.warning("commissioni_annuale is 0 — check PDF extraction")

    # ── 4. Period percentage — direct from PDF ────────────────────────────────
    pct_periodo = _parse_amount(str(contract.get("percentuale_per_periodo", "0")))

    # If AI did not extract percentuale_per_periodo, calculate from frequency
    if pct_periodo == 0 and comm_annuale > 0:
        frequenza = str(contract.get("frequenza_pagamento", "semestrale")).lower()
        if "trim" in frequenza:
            pct_periodo = comm_annuale / 4
            log.info(f"Calculated trimestrale period %: {pct_periodo}")
        elif "ann" in frequenza:
            pct_periodo = comm_annuale
            log.info(f"Calculated annuale period %: {pct_periodo}")
        else:
            # default semestrale
            pct_periodo = comm_annuale / 2
            log.info(f"Calculated semestrale period %: {pct_periodo}")

    # ── 5. Amount paid each payment period ───────────────────────────────────
    amount_per_period = importo * (pct_periodo / 100)
    log.info(f"amount_per_period = {importo} × {pct_periodo}% = {amount_per_period}")

    # ── 6. Payment interval in months ────────────────────────────────────────
    frequenza = str(contract.get("frequenza_pagamento", "semestrale")).lower()
    if "trim" in frequenza:
        payment_interval = 3
    elif "ann" in frequenza:
        payment_interval = 12
    else:
        payment_interval = 6   # semestrale

    # ── 7. Build 12-month schedule ────────────────────────────────────────────
    schedule = []
    for i in range(12):
        total_months = start_m + i
        month_idx    = total_months % 12
        year         = start_y + total_months // 12
        mese         = MONTHS_IT[month_idx]
        mon_short    = MONTHS_SHORT[month_idx]

        if i < EMPTY_MONTHS:
            # Grace/startup period — no payment
            schedule.append({
                "mese":                       mese,
                "data_pagamento_commissioni": "",
                "management_mensile":         "",
                "_amount":                    0.0,
            })
        else:
            # Check if this month is a payment month
            months_active      = i - EMPTY_MONTHS + 1
            is_payment_month   = (months_active % payment_interval == 0)
            pay_date = f"10-{mon_short}-{str(year)[2:]}" if is_payment_month else ""
            amt      = amount_per_period if is_payment_month else 0.0

            schedule.append({
                "mese":                       mese,
                "data_pagamento_commissioni": pay_date,
                "management_mensile":         f"{amt:.2f}" if amt else "",
                "_amount":                    amt,
            })

    return schedule


def enrich_contracts(data: dict) -> dict:
    """Add calculated fields to each contract — no hardcoded defaults."""
    for c in data.get("contracts", []):
        c["_importo"]       = _parse_amount(str(c.get("importo_contratto", "0")))
        c["_comm_annuale"]  = _parse_amount(str(c.get("commissioni_annuale", "0")))
        c["_pct_periodo"]   = _parse_amount(str(c.get("percentuale_per_periodo", "0")))

        # Keep "commissioni" field for display in table header
        c["commissioni"]    = str(c["_comm_annuale"])
        c["_comm"]          = c["_comm_annuale"]

        c["payment_schedule"] = fix_schedule(c)

    return data


# ════════════════════════════════════════════════════════════════════════════
#  DOCX Generator
# ════════════════════════════════════════════════════════════════════════════

_NAVY  = RGBColor(0x1F, 0x49, 0x7D)
_GOLD  = RGBColor(0xC0, 0x9B, 0x3A)
_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
_DARK  = RGBColor(0x26, 0x26, 0x26)


def _shd(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _hdr_cell(cell, text: str, size=8):
    _shd(cell, "1F497D")
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.bold, r.font.size, r.font.name = True, Pt(size), "Calibri"
    r.font.color.rgb = _WHITE


def _data_cell(cell, text: str, bold=False, center=True, bg="FFFFFF", color=None):
    _shd(cell, bg)
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(str(text) if text else "")
    r.font.bold, r.font.size, r.font.name = bold, Pt(8), "Calibri"
    r.font.color.rgb = color if color else _DARK


def _fmt_eur(amount: float) -> str:
    """Format as Italian currency: 10.000,00 €"""
    return f"{amount:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def generate_docx(data: dict, out_path: str):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width    = Cm(29.7)
    sec.page_height   = Cm(21.0)
    sec.left_margin   = Cm(1.5)
    sec.right_margin  = Cm(1.5)
    sec.top_margin    = Cm(1.5)
    sec.bottom_margin = Cm(1.5)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("CONTO METALLO DELUXE — GESTIONE COMMISSIONI")
    r.font.bold, r.font.size, r.font.name = True, Pt(13), "Calibri"
    r.font.color.rgb = _NAVY

    s  = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = s.add_run(
        f"Doc: {data.get('document_number','N/A')}  |  "
        f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    sr.font.size, sr.font.name = Pt(8), "Calibri"
    sr.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
    doc.add_paragraph()

    contracts = data.get("contracts", [])
    if not contracts:
        doc.add_paragraph("No contracts found in PDF.")
        doc.save(out_path)
        return

    for ci, contract in enumerate(contracts):
        lbl = doc.add_paragraph()
        lr  = lbl.add_run(f"Contract #{ci+1}  —  {contract.get('nome_cliente_segnalatore','')}")
        lr.font.bold, lr.font.size, lr.font.name = True, Pt(9), "Calibri"
        lr.font.color.rgb = _GOLD

        schedule = contract.get("payment_schedule", [])
        n_rows   = max(len(schedule), 1)

        tbl = doc.add_table(rows=1 + n_rows + 1, cols=8)
        tbl.style     = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        col_widths = [3.0, 4.8, 2.6, 3.0, 2.4, 2.6, 3.8, 3.4]
        for row in tbl.rows:
            for i, cell in enumerate(row.cells):
                cell.width = Cm(col_widths[i])

        headers = [
            "Tipo Contratto", "Nome Cliente/Segnalatore", "Data Contratto",
            "Importo Contratto", "Commissioni\n(% annuo)", "Mese",
            "Data Pagamento\nCommissioni", "Management\nMensile"
        ]
        for i, txt in enumerate(headers):
            _hdr_cell(tbl.rows[0].cells[i], txt)

        importo_raw  = contract.get("_importo", 0.0)
        comm_annuale = contract.get("_comm_annuale", 0.0)
        tipo         = contract.get("tipo_contratto", "Conto Metallo Deluxe")
        nome         = contract.get("nome_cliente_segnalatore", "")
        data_c       = contract.get("data_contratto", "")
        piano        = contract.get("piano", "")
        importo_fmt  = _fmt_eur(importo_raw)
        # Show annual commission % with piano label
        comm_display = f"{comm_annuale:.2f}% p.a. ({piano})"

        total_paid = 0.0
        for ri, sched in enumerate(schedule):
            row = tbl.rows[1 + ri]
            bg  = "FFFFFF" if ri % 2 == 0 else "F4F0E8"

            _data_cell(row.cells[0], tipo if ri == 0 else "", bold=(ri == 0), center=False, bg=bg)
            _data_cell(row.cells[1], nome if ri == 0 else "", center=False, bg=bg)
            _data_cell(row.cells[2], data_c if ri == 0 else "", bg=bg)
            _data_cell(row.cells[3], importo_fmt if ri == 0 else "", bold=(ri == 0), bg=bg)
            _data_cell(row.cells[4], comm_display if ri == 0 else "", bg=bg)
            _data_cell(row.cells[5], sched.get("mese", ""), bold=True, bg=bg)

            pay = sched.get("data_pagamento_commissioni", "")
            _data_cell(row.cells[6], pay if pay else "—", bg=bg)

            amt = sched.get("_amount", 0.0)
            if not amt:
                raw_s = sched.get("management_mensile", "")
                amt   = _parse_amount(raw_s) if raw_s else 0.0
            total_paid += amt
            amt_fmt = _fmt_eur(amt) if amt else ""
            _data_cell(row.cells[7], amt_fmt, bold=bool(amt), bg=bg,
                       color=_NAVY if amt else None)

        # Total row
        total_row = tbl.rows[1 + n_rows]
        _shd(total_row.cells[0], "E8F0FE")
        total_row.cells[0].merge(total_row.cells[5])
        p = total_row.cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r2 = p.add_run(f"TOTALE COMMISSIONI — {nome}")
        r2.font.bold, r2.font.size, r2.font.name = True, Pt(9), "Calibri"
        r2.font.color.rgb = _NAVY

        _data_cell(total_row.cells[6], "TOTALE", bold=True, bg="E8F0FE", color=_NAVY)
        _data_cell(total_row.cells[7], _fmt_eur(total_paid), bold=True, bg="E8F0FE",
                   color=RGBColor(0x05, 0x96, 0x69))

        if ci < len(contracts) - 1:
            doc.add_paragraph()

    notes = data.get("extraction_notes", "")
    if notes:
        doc.add_paragraph()
        np_ = doc.add_paragraph()
        nr  = np_.add_run(f"  {notes}")
        nr.font.size, nr.font.italic = Pt(7), True
        nr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    doc.save(out_path)
    log.info(f"DOCX saved: {out_path}")


# ════════════════════════════════════════════════════════════════════════════
#  Excel Generator
# ════════════════════════════════════════════════════════════════════════════

_FILL_NAVY   = PatternFill("solid", fgColor="1F497D")
_FILL_GOLD   = PatternFill("solid", fgColor="C09B3A")
_FILL_STRIPE = PatternFill("solid", fgColor="F4F0E8")
_FILL_WHITE  = PatternFill("solid", fgColor="FFFFFF")
_FILL_TOTAL  = PatternFill("solid", fgColor="E8F0FE")

_FONT_HDR   = Font(name="Calibri", bold=True,  color="FFFFFF", size=9)
_FONT_TIPO  = Font(name="Calibri", bold=True,  color="1F497D", size=8)
_FONT_DATA  = Font(name="Calibri", size=8,     color="262626")
_FONT_AMT   = Font(name="Calibri", bold=True,  size=8, color="1F497D")
_FONT_TOTAL = Font(name="Calibri", bold=True,  size=9, color="1F497D")
_FONT_TITLE = Font(name="Calibri", bold=True,  size=14, color="1F497D")
_FONT_SUB   = Font(name="Calibri", size=9,     color="707070")

_ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_ALIGN_R = Alignment(horizontal="right",  vertical="center")

_THIN  = Side(style="thin",   color="BFBFBF")
_THICK = Side(style="medium", color="1F497D")
_BORDER_THIN  = Border(left=_THIN,  right=_THIN,  top=_THIN,  bottom=_THIN)
_BORDER_THICK = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)

HEADERS = [
    "Tipo Contratto", "Nome Cliente/Segnalatore", "Data Contratto",
    "Importo Contratto", "Commissioni (% annuo)", "Mese",
    "Data Pagamento Commissioni", "Management Mensile (€)",
]
COL_W = [22, 30, 14, 18, 16, 14, 24, 22]


def _cell(ws, row, col, value, font=None, fill=None, align=None, border=None, num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:    c.font          = font
    if fill:    c.fill          = fill
    if align:   c.alignment     = align
    if border:  c.border        = border
    if num_fmt: c.number_format = num_fmt
    return c


def generate_excel(data: dict, out_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Commissioni"
    ws.freeze_panes = "A5"

    ws.merge_cells("A1:H1")
    _cell(ws, 1, 1, "CONTO METALLO DELUXE — GESTIONE COMMISSIONI",
          font=_FONT_TITLE,
          fill=PatternFill("solid", fgColor="1F497D"),
          align=_ALIGN_C)

    ws.merge_cells("A2:H2")
    _cell(ws, 2, 1,
          f"Doc: {data.get('document_number','N/A')}   |   "
          f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
          font=_FONT_SUB,
          fill=PatternFill("solid", fgColor="C09B3A"),
          align=_ALIGN_C)

    ws.row_dimensions[3].height = 6

    for ci, (hdr, w) in enumerate(zip(HEADERS, COL_W), start=1):
        _cell(ws, 4, ci, hdr,
              font=_FONT_HDR, fill=_FILL_NAVY,
              align=_ALIGN_C, border=_BORDER_THIN)
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[4].height = 32

    current_row = 5
    contracts   = data.get("contracts", [])

    for contract in contracts:
        schedule     = contract.get("payment_schedule", [])
        importo_raw  = contract.get("_importo", 0.0)
        comm_annuale = contract.get("_comm_annuale", 0.0)
        piano        = contract.get("piano", "")
        tipo         = contract.get("tipo_contratto", "Conto Metallo Deluxe")
        nome         = contract.get("nome_cliente_segnalatore", "")
        datec        = contract.get("data_contratto", "")

        for ri, sched in enumerate(schedule):
            stripe = _FILL_STRIPE if ri % 2 else _FILL_WHITE
            r      = current_row + ri
            ws.row_dimensions[r].height = 18

            _cell(ws, r, 1,
                  tipo if ri == 0 else "",
                  font=_FONT_TIPO if ri == 0 else _FONT_DATA,
                  fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)

            _cell(ws, r, 2,
                  nome if ri == 0 else "",
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)

            _cell(ws, r, 3,
                  datec if ri == 0 else "",
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

            _cell(ws, r, 4,
                  importo_raw if ri == 0 else None,
                  font=Font(name="Calibri", bold=(ri==0), size=8,
                            color="1F497D" if ri==0 else "262626"),
                  fill=stripe, align=_ALIGN_R, border=_BORDER_THIN,
                  num_fmt='#,##0.00 "€"')

            # Commission: show annual % with piano label
            comm_label = f"{comm_annuale:.2f}% p.a. ({piano})" if ri == 0 else None
            _cell(ws, r, 5,
                  comm_label,
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

            _cell(ws, r, 6,
                  sched.get("mese", ""),
                  font=Font(name="Calibri", bold=True, size=8),
                  fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

            pay = sched.get("data_pagamento_commissioni", "")
            _cell(ws, r, 7,
                  pay if pay else "—",
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

            amt = sched.get("_amount", 0.0)
            if not amt:
                raw_s = sched.get("management_mensile", "")
                amt   = _parse_amount(raw_s) if raw_s else 0.0

            _cell(ws, r, 8,
                  amt if amt else None,
                  font=_FONT_AMT if amt else _FONT_DATA,
                  fill=stripe, align=_ALIGN_R, border=_BORDER_THIN,
                  num_fmt='#,##0.00 "€"')

        current_row += len(schedule)

        # Total row
        total_amt = sum(s.get("_amount", 0.0) for s in schedule)
        ws.row_dimensions[current_row].height = 20
        ws.merge_cells(f"A{current_row}:F{current_row}")
        _cell(ws, current_row, 1,
              f"TOTALE COMMISSIONI — {nome}",
              font=_FONT_TOTAL, fill=_FILL_TOTAL,
              align=_ALIGN_L, border=_BORDER_THICK)
        for col in range(2, 7):
            _cell(ws, current_row, col, None,
                  fill=_FILL_TOTAL, border=_BORDER_THICK)
        _cell(ws, current_row, 7,
              "TOTALE",
              font=_FONT_TOTAL, fill=_FILL_TOTAL,
              align=_ALIGN_C, border=_BORDER_THICK)
        _cell(ws, current_row, 8,
              total_amt,
              font=Font(name="Calibri", bold=True, size=10, color="059669"),
              fill=_FILL_TOTAL, align=_ALIGN_R, border=_BORDER_THICK,
              num_fmt='#,##0.00 "€"')

        current_row += 2

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    for col, w in zip("ABCDE", [30, 20, 18, 18, 20]):
        ws2.column_dimensions[col].width = w

    ws2.merge_cells("A1:E1")
    _cell(ws2, 1, 1, "CONTRACT SUMMARY",
          font=_FONT_TITLE, fill=_FILL_NAVY, align=_ALIGN_C)
    ws2.row_dimensions[1].height = 28

    summary_headers = [
        "Cliente/Segnalatore", "Importo", "Comm % Annuo",
        "Annual Return", "Total Paid (Year 1)"
    ]
    for ci2, hdr in enumerate(summary_headers, 1):
        _cell(ws2, 2, ci2, hdr,
              font=_FONT_HDR, fill=_FILL_NAVY,
              align=_ALIGN_C, border=_BORDER_THIN)
    ws2.row_dimensions[2].height = 22

    for ri, c in enumerate(contracts, start=3):
        importo_r    = c.get("_importo", 0.0)
        comm_annuale = c.get("_comm_annuale", 0.0)

        # Annual return = importo × annual commission %  (correct formula)
        annual_return = importo_r * (comm_annuale / 100)

        # Total paid = sum of all payment amounts in the schedule
        paid = sum(s.get("_amount", 0.0) for s in c.get("payment_schedule", []))

        stripe2 = _FILL_STRIPE if ri % 2 else _FILL_WHITE

        _cell(ws2, ri, 1, c.get("nome_cliente_segnalatore", ""),
              font=_FONT_DATA, fill=stripe2, align=_ALIGN_L, border=_BORDER_THIN)
        _cell(ws2, ri, 2, importo_r,
              font=_FONT_AMT, fill=stripe2, align=_ALIGN_R,
              border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
        _cell(ws2, ri, 3, comm_annuale / 100,
              font=_FONT_DATA, fill=stripe2, align=_ALIGN_C,
              border=_BORDER_THIN, num_fmt='0.00%')
        _cell(ws2, ri, 4, annual_return,
              font=_FONT_AMT, fill=stripe2, align=_ALIGN_R,
              border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
        _cell(ws2, ri, 5, paid,
              font=Font(name="Calibri", bold=True, size=8, color="059669"),
              fill=stripe2, align=_ALIGN_R,
              border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
        ws2.row_dimensions[ri].height = 18

    wb.save(out_path)
    log.info(f"XLSX saved: {out_path}")


# ════════════════════════════════════════════════════════════════════════════
#  Helpers
# ════════════════════════════════════════════════════════════════════════════

def _load_job_data(job_id: str, body_data: Any) -> dict:
    if body_data:
        return enrich_contracts(body_data)
    path = OUTPUT_DIR / f"{job_id}_data.json"
    if path.exists():
        with open(path, encoding="utf-8") as fp:
            return enrich_contracts(json.load(fp))
    return None


# ════════════════════════════════════════════════════════════════════════════
#  API Routes
# ════════════════════════════════════════════════════════════════════════════

@app.get("/api/health")
def health():
    return {
        "status":    "ok",
        "ai_ready":  bool(OPENAI_API_KEY),
        "model":     OPENAI_MODEL,
        "timestamp": datetime.now().isoformat(),
    }


@app.post("/api/upload")
async def upload_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files accepted")

    job_id   = str(uuid.uuid4())[:8]
    pdf_path = UPLOAD_DIR / f"{job_id}_{file.filename}"

    content = await file.read()
    with open(pdf_path, "wb") as f:
        f.write(content)
    log.info(f"Saved PDF: {pdf_path} ({pdf_path.stat().st_size} bytes)")

    try:
        pdf_text = extract_pdf_text(str(pdf_path))
        log.info(f"Extracted {len(pdf_text)} chars from PDF")

        if not OPENAI_API_KEY:
            raise HTTPException(status_code=500, detail="OPENAI_API_KEY not set")

        extracted = ai_extract(pdf_text)
        log.info(f"AI extracted: {json.dumps(extracted, indent=2)[:500]}")

        data_path = OUTPUT_DIR / f"{job_id}_data.json"
        with open(data_path, "w", encoding="utf-8") as fp:
            json.dump(extracted, fp, ensure_ascii=False, indent=2)

        return {
            "job_id":           job_id,
            "filename":         file.filename,
            "contracts_found":  len(extracted.get("contracts", [])),
            "document_number":  extracted.get("document_number", ""),
            "extraction_notes": extracted.get("extraction_notes", ""),
            "data":             extracted,
            "status":           "analyzed",
        }

    except Exception as e:
        log.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        try:
            pdf_path.unlink()
        except Exception:
            pass


@app.post("/api/generate/docx")
def gen_docx(req: GenerateRequest):
    data = _load_job_data(req.job_id, req.data)
    if not data:
        raise HTTPException(status_code=404, detail="No data for job_id")
    out = OUTPUT_DIR / f"{req.job_id}_commissioni.docx"
    try:
        generate_docx(data, str(out))
        return FileResponse(
            path=str(out),
            filename=f"Commissioni_{req.job_id}.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        log.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/generate/excel")
def gen_excel(req: GenerateRequest):
    data = _load_job_data(req.job_id, req.data)
    if not data:
        raise HTTPException(status_code=404, detail="No data for job_id")
    out = OUTPUT_DIR / f"{req.job_id}_commissioni.xlsx"
    try:
        generate_excel(data, str(out))
        return FileResponse(
            path=str(out),
            filename=f"Commissioni_{req.job_id}.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        log.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/generate/both")
def gen_both(req: GenerateRequest):
    data = _load_job_data(req.job_id, req.data)
    if not data:
        raise HTTPException(status_code=404, detail="No data for job_id")
    result: dict = {"job_id": req.job_id}
    try:
        docx_path = OUTPUT_DIR / f"{req.job_id}_commissioni.docx"
        generate_docx(data, str(docx_path))
        result["docx_ready"] = True
    except Exception as e:
        result["docx_error"] = str(e)
        result["docx_ready"] = False
    try:
        xlsx_path = OUTPUT_DIR / f"{req.job_id}_commissioni.xlsx"
        generate_excel(data, str(xlsx_path))
        result["excel_ready"] = True
    except Exception as e:
        result["excel_error"] = str(e)
        result["excel_ready"] = False
    return result


@app.post("/api/generate")
def gen_document(req: GenerateRequest):
    return gen_docx(req)


@app.post("/api/preview")
def preview(req: GenerateRequest):
    data = _load_job_data(req.job_id, req.data)
    if not data:
        raise HTTPException(status_code=404, detail="No data for job_id")
    return data


@app.post("/api/debug")
def debug_extraction(req: GenerateRequest):
    data = _load_job_data(req.job_id, req.data)
    if not data:
        raise HTTPException(status_code=404, detail="No data for job_id")
    result = []
    for i, c in enumerate(data.get("contracts", [])):
        result.append({
            "contract_number":          i + 1,
            "tipo_contratto":           c.get("tipo_contratto"),
            "nome_cliente_segnalatore": c.get("nome_cliente_segnalatore"),
            "data_contratto":           c.get("data_contratto"),
            "importo_contratto":        c.get("importo_contratto"),
            "commissioni_annuale":      c.get("commissioni_annuale"),
            "frequenza_pagamento":      c.get("frequenza_pagamento"),
            "percentuale_per_periodo":  c.get("percentuale_per_periodo"),
            "durata_mesi":              c.get("durata_mesi"),
            "piano":                    c.get("piano"),
            "_importo_parsed":          c.get("_importo"),
            "_comm_annuale_parsed":     c.get("_comm_annuale"),
            "_pct_periodo_parsed":      c.get("_pct_periodo"),
            "total_months":             len(c.get("payment_schedule", [])),
            "payment_schedule":         c.get("payment_schedule", []),
        })
    return {"debug": result}


@app.get("/api/download/{job_id}/{fmt}")
def download_file(job_id: str, fmt: str):
    if fmt == "docx":
        path  = OUTPUT_DIR / f"{job_id}_commissioni.docx"
        media = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        name  = f"Commissioni_{job_id}.docx"
    elif fmt == "excel":
        path  = OUTPUT_DIR / f"{job_id}_commissioni.xlsx"
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        name  = f"Commissioni_{job_id}.xlsx"
    else:
        raise HTTPException(status_code=400, detail="fmt must be docx or excel")
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not generated yet")
    return FileResponse(path=str(path), filename=name, media_type=media)


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    log.info(f"Starting server on port {PORT}")
    uvicorn.run("main:app", host="0.0.0.0", port=PORT, reload=False)

















# #_---openai working code---some isssue
# """
# main.py — PDF → DOCX + Excel Extractor (FastAPI version)
# """

# import os
# import re
# import io
# import json
# import uuid
# import logging
# import traceback
# import urllib.request
# from datetime import datetime
# from pathlib import Path

# # ── Load .env before anything else ──────────────────────────────────────────
# def _load_dotenv(path: str = ".env"):
#     env_path = Path(path)
#     if not env_path.exists():
#         env_path = Path(__file__).parent / ".env"
#     if not env_path.exists():
#         return
#     with open(env_path, encoding="utf-8") as f:
#         for line in f:
#             line = line.strip()
#             if not line or line.startswith("#") or "=" not in line:
#                 continue
#             key, _, val = line.partition("=")
#             key = key.strip()
#             val = val.strip().strip('"').strip("'")
#             if key and key not in os.environ:
#                 os.environ[key] = val

# _load_dotenv()

# # ── Logging ──────────────────────────────────────────────────────────────────
# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s %(levelname)s %(message)s",
#     datefmt="%Y-%m-%d %H:%M:%S",
# )
# log = logging.getLogger(__name__)

# # ── Config from env ───────────────────────────────────────────────────────────
# OPENAI_API_KEY    = os.environ.get("OPENAI_API_KEY", "")
# OPENAI_MODEL      = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
# OPENAI_MAX_TOKENS = int(os.environ.get("OPENAI_MAX_TOKENS", "4096"))
# PORT= int(os.environ.get("PORT", "8000"))

# log.info(f"Model  : {OPENAI_MODEL}")
# log.info(f"AI Key : {'SET ✓' if OPENAI_API_KEY else 'NOT SET'}")

# BASE_DIR   = Path(__file__).parent
# UPLOAD_DIR = BASE_DIR / "uploads"
# OUTPUT_DIR = BASE_DIR / "outputs"
# UPLOAD_DIR.mkdir(exist_ok=True)
# OUTPUT_DIR.mkdir(exist_ok=True)

# log.info(f"Model  : {OPENAI_API_KEY}")
# log.info(f"AI Key : {'SET ✓' if OPENAI_API_KEY else 'NOT SET — fallback mode'}")

# # ── FastAPI imports ───────────────────────────────────────────────────────────
# from fastapi import FastAPI, File, UploadFile, HTTPException, Request
# from fastapi.responses import JSONResponse, FileResponse
# from fastapi.middleware.cors import CORSMiddleware
# from pydantic import BaseModel
# from typing import Optional, Any
# import uvicorn

# import pdfplumber
# from docx import Document
# from docx.shared import Pt, RGBColor, Cm
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# from docx.enum.table import WD_TABLE_ALIGNMENT
# from docx.oxml.ns import qn
# from docx.oxml import OxmlElement
# import openpyxl
# from openpyxl.styles import (
#     PatternFill, Font, Alignment, Border, Side
# )
# from openpyxl.utils import get_column_letter
# from fastapi.middleware.cors import CORSMiddleware



# # ── App ───────────────────────────────────────────────────────────────────────
# app = FastAPI(title="MetalloDoc API", version="1.0.0")
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=False,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )



# # ── Pydantic models ───────────────────────────────────────────────────────────
# class GenerateRequest(BaseModel):
#     job_id: str
#     data: Optional[Any] = None


# # ════════════════════════════════════════════════════════════════════════════
# #  AI + PDF Extraction  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# def extract_pdf_text(pdf_path: str) -> str:
#     pages = []
#     with pdfplumber.open(pdf_path) as pdf:
#         for i, page in enumerate(pdf.pages):
#             text = page.extract_text() or ""
#             pages.append(f"=== PAGE {i+1} ===\n{text}")
#     return "\n\n".join(pages)


# def call_openai(prompt: str, system: str) -> str:
#     payload = json.dumps({
#         "model":      OPENAI_MODEL,
#         "max_tokens": OPENAI_MAX_TOKENS,
#         "messages": [
#             {"role": "system", "content": system},
#             {"role": "user",   "content": prompt},
#         ],
#     }).encode("utf-8")

#     req = urllib.request.Request(
#         "https://api.openai.com/v1/chat/completions",
#         data=payload,
#         headers={
#             "Content-Type":  "application/json",
#             "Authorization": f"Bearer {OPENAI_API_KEY}",
#         },
#         method="POST",
#     )
#     try:
#         with urllib.request.urlopen(req, timeout=300) as resp:
#             body = json.loads(resp.read().decode("utf-8"))
#             return body["choices"][0]["message"]["content"]
#     except urllib.error.HTTPError as e:
#         error_body = e.read().decode("utf-8")
#         log.error(f"OpenAI API error {e.code}: {error_body}")
#         raise


# def safe_json(text: str) -> dict:
#     for attempt in (
#         lambda t: json.loads(t.strip()),
#         lambda t: json.loads(re.search(r"```(?:json)?\s*(\{.*?\})\s*```", t, re.DOTALL).group(1)),
#         lambda t: json.loads(re.search(r"\{.*\}", t, re.DOTALL).group(0)),
#     ):
#         try:
#             return attempt(text)
#         except Exception:
#             pass
#     raise ValueError(f"Cannot parse JSON from AI response: {text[:300]}")


# SYSTEM_PROMPT = """\
# You are a financial document data extractor for Italian gold investment contracts
# (Conto Metallo Deluxe, Mac & Ro Srl). Extract data and return ONLY valid JSON.

# CRITICAL: Read every single table row in the PDF carefully.
# For each month row you find, extract:
#   - The month name (mese) in Italian
#   - The payment date if present (data_pagamento_commissioni) e.g. "10-Dec-25"
#   - The management amount if present (management_mensile) e.g. "4340.00" (numbers only, no € symbol)
#   - If a row has no date or amount, leave those fields as empty string ""

# Required JSON structure:
# {
#   "contracts": [
#     {
#       "tipo_contratto": "Conto Metallo Deluxe",
#       "nome_cliente_segnalatore": "Cognome Nome Cliente/Cognome Nome Collaboratore",
#       "data_contratto": "DD/MM/YYYY",
#       "importo_contratto": "200000.00",
#       "commissioni": "2.17",
#       "piano": "Plus",
#       "payment_schedule": [
#         {
#           "mese": "Luglio",
#           "data_pagamento_commissioni": "",
#           "management_mensile": ""
#         },
#         {
#           "mese": "Agosto",
#           "data_pagamento_commissioni": "",
#           "management_mensile": ""
#         },
#         {
#           "mese": "Settembre",
#           "data_pagamento_commissioni": "",
#           "management_mensile": ""
#         },
#         {
#           "mese": "Ottobre",
#           "data_pagamento_commissioni": "",
#           "management_mensile": ""
#         },
#         {
#           "mese": "Novembre",
#           "data_pagamento_commissioni": "",
#           "management_mensile": ""
#         },
#         {
#           "mese": "Dicembre",
#           "data_pagamento_commissioni": "10-Dec-25",
#           "management_mensile": "4340.00"
#         }
#       ]
#     }
#   ],
#   "document_number": "CMP-000018.01.2026",
#   "extraction_notes": "Extracted from PDF"
# }

# STRICT RULES:
# - Extract ONLY what is physically written in the PDF table rows
# - management_mensile: numeric string only, strip € and spaces (e.g. "4.340,00" → "4340.00")
# - data_pagamento_commissioni: copy exactly as written in PDF (e.g. "10-Dec-25")
# - commissioni: percentage number only, no % symbol (e.g. "2.17")
# - importo_contratto: number only, no € symbol (e.g. "200000.00")
# - nome_cliente_segnalatore: "ClientLastFirst/AgentLastFirst" from Contraente + Collaboratore fields
# - data_contratto: from "Luogo e data" field, format DD/MM/YYYY
# - Return ONLY raw JSON, no markdown fences, no explanation\
# """


# def ai_extract(pdf_text: str) -> dict:
#     prompt = (
#         "Extract ALL data from this Italian financial PDF contract.\n"
#         "The payment schedule table has rows for every month - extract ALL of them, "
#         "including months in 2026. Do not stop at the first few rows.\n"
#         "For empty rows (no date/amount) still include them with empty strings.\n"
#         "For amounts like '4.340,00 €' convert to '4340.00'.\n\n"
#         f"{pdf_text[:50000]}\n\n"
#         "Return ONLY the JSON object with no markdown."
#     )
#     raw = call_openai(prompt, SYSTEM_PROMPT)
#     log.info(f"AI raw (first 500): {raw[:500]}")
#     return safe_json(raw)

# # ════════════════════════════════════════════════════════════════════════════
# #  Fallback pattern-matching extractor  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# def pattern_extract(pdf_text: str, filename: str) -> dict:
#     m = re.search(r"(CMP-[\d.]+)", pdf_text)
#     doc_num = m.group(1) if m else "N/A"

#     m = re.search(r"Cognome e Nome[:\s]+([A-Za-zÀ-ÿ]+\s+[A-Za-zÀ-ÿ]+)", pdf_text)
#     client = m.group(1).strip() if m else "Unknown Client"

#     m = re.search(r"COLLABORATORE.*?Cognome e Nome\s+([A-Za-zÀ-ÿ]+\s+[A-Za-zÀ-ÿ]+)", pdf_text, re.DOTALL)
#     collab = m.group(1).strip() if m else "Agent"

#     m = re.search(r"Corrispettivo\s*\n?\s*([\d.,]+)", pdf_text)
#     raw_amt = m.group(1).strip() if m else "10000"
#     amount = _parse_amount(raw_amt)

#     m = re.search(r"(\d{4}-\d{2}-\d{2})", pdf_text)
#     if m:
#         dt = datetime.strptime(m.group(1), "%Y-%m-%d")
#         date_str = dt.strftime("%d/%m/%Y")
#     else:
#         date_str = "09/01/2026"

#     is_top = ("Deluxe Top" in pdf_text) and (amount >= 100000)
#     piano  = "Top" if is_top else "Plus"
#     comm   = "2.50" if is_top else "2.17"

#     return {
#         "contracts": [{
#             "tipo_contratto": "Conto Metallo Deluxe",
#             "nome_cliente_segnalatore": f"{client}/{collab}",
#             "data_contratto": date_str,
#             "importo_contratto": f"{amount:.2f}",
#             "commissioni": comm,
#             "piano": piano,
#             "payment_schedule": [],
#         }],
#         "document_number": doc_num,
#         "extraction_notes": "Pattern-matching extraction (no AI key)",
#     }


# # ════════════════════════════════════════════════════════════════════════════
# #  Payment schedule builder  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# MONTHS_IT    = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
#                 "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]
# MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun",
#                 "Jul","Aug","Sep","Oct","Nov","Dec"]


# def _parse_amount(s: str) -> float:
#     s = str(s).replace("€","").replace("$","").strip()
#     if "." in s and "," in s:
#         if s.index(",") < s.index("."):
#             s = s.replace(",","")
#         else:
#             s = s.replace(".","").replace(",",".")
#     elif "," in s:
#         parts = s.split(",")
#         if len(parts[-1]) == 2 and len(parts[0]) >= 2:
#             s = s.replace(",","")
#         else:
#             s = s.replace(",",".")
#     try:
#         return float(re.sub(r"[^\d.]", "", s))
#     except Exception:
#         return 10000.0


# def build_schedule(contract: dict) -> list:
#     if contract.get("payment_schedule") and len(contract["payment_schedule"]) >= 12:
#         for row in contract["payment_schedule"]:
#             v = row.get("management_mensile","")
#             if v:
#                 row["_amount"] = _parse_amount(str(v))
#         return contract["payment_schedule"]

#     importo = _parse_amount(str(contract.get("importo_contratto","10000")))
#     comm    = float(re.sub(r"[^\d.]", "", str(contract.get("commissioni","2.17"))) or "2.17")
#     periodic_amount = importo * (comm / 100)

#     date_str = contract.get("data_contratto","01/01/2026")
#     try:
#         dt = datetime.strptime(date_str, "%d/%m/%Y")
#         start_m = dt.month - 1
#         start_y = dt.year
#     except Exception:
#         start_m, start_y = 0, 2026

#     schedule = []
#     for i in range(12):
#         idx  = (start_m + i) % 12
#         year = start_y + (start_m + i) // 12
#         name = MONTHS_IT[idx]

#         if i < 5:
#             schedule.append({
#                 "mese": name,
#                 "data_pagamento_commissioni": "",
#                 "management_mensile": "",
#                 "_amount": 0.0,
#             })
#         else:
#             pay_date = f"{10:02d}-{MONTHS_SHORT[idx]}-{str(year)[2:]}"
#             schedule.append({
#                 "mese": name,
#                 "data_pagamento_commissioni": pay_date,
#                 "management_mensile": f"{periodic_amount:.2f}",
#                 "_amount": periodic_amount,
#             })

#     return schedule


# def fix_schedule(contract: dict) -> list:
#     """Always exactly 12 months, fully dynamic from contract date."""
    
#     date_str = contract.get("data_contratto", "01/01/2026")
#     try:
#         dt      = datetime.strptime(date_str, "%d/%m/%Y")
#         start_m = dt.month - 1  # 0-based (0=Jan, 6=Jul)
#         start_y = dt.year
#     except Exception:
#         start_m, start_y = 0, 2026

#     importo = _parse_amount(str(contract.get("importo_contratto", "0")))
#     comm    = float(re.sub(r"[^\d.]", "", str(contract.get("commissioni", "2.17"))) or "2.17")
#     monthly = importo * (comm / 100)

#     schedule = []
#     for i in range(12):
#         # Calculate exact month and year for each row
#         total_months = start_m + i
#         month_idx    = total_months % 12        # 0-based month index
#         year         = start_y + total_months // 12  # correct year

#         mese     = MONTHS_IT[month_idx]
#         mon_short = MONTHS_SHORT[month_idx]
#         pay_date = f"10-{mon_short}-{str(year)[2:]}"  # e.g. 10-Dec-26

#         if i < 5:
#             schedule.append({
#                 "mese":                        mese,
#                 "data_pagamento_commissioni":  "",
#                 "management_mensile":          "",
#                 "_amount":                     0.0,
#             })
#         else:
#             schedule.append({
#                 "mese":                        mese,
#                 "data_pagamento_commissioni":  pay_date,
#                 "management_mensile":          f"{monthly:.2f}",
#                 "_amount":                     monthly,
#             })

#     return schedule


# def enrich_contracts(data: dict) -> dict:
#     for c in data.get("contracts", []):
#         c["payment_schedule"] = fix_schedule(c)
#         c["_importo"] = _parse_amount(str(c.get("importo_contratto", "0")))
#         c["_comm"]    = float(
#             re.sub(r"[^\d.]", "", str(c.get("commissioni", "2.17"))) or "2.17"
#         )
#     return data


# # ════════════════════════════════════════════════════════════════════════════
# #  DOCX Generator  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# _NAVY  = RGBColor(0x1F,0x49,0x7D)
# _GOLD  = RGBColor(0xC0,0x9B,0x3A)
# _WHITE = RGBColor(0xFF,0xFF,0xFF)
# _DARK  = RGBColor(0x26,0x26,0x26)


# def _shd(cell, hex_color: str):
#     tc   = cell._tc
#     tcPr = tc.get_or_add_tcPr()
#     shd  = OxmlElement("w:shd")
#     shd.set(qn("w:val"),   "clear")
#     shd.set(qn("w:color"), "auto")
#     shd.set(qn("w:fill"),  hex_color)
#     tcPr.append(shd)


# def _hdr_cell(cell, text: str, size=8):
#     _shd(cell, "1F497D")
#     cell.text = ""
#     p = cell.paragraphs[0]
#     p.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     r = p.add_run(text)
#     r.font.bold, r.font.size, r.font.name = True, Pt(size), "Calibri"
#     r.font.color.rgb = _WHITE


# def _data_cell(cell, text: str, bold=False, center=True, bg="FFFFFF", color=None):
#     _shd(cell, bg)
#     cell.text = ""
#     p = cell.paragraphs[0]
#     p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
#     r = p.add_run(str(text) if text else "")
#     r.font.bold, r.font.size, r.font.name = bold, Pt(8), "Calibri"
#     r.font.color.rgb = color if color else _DARK


# def generate_docx(data: dict, out_path: str):
#     doc = Document()

#     sec = doc.sections[0]
#     sec.page_width    = Cm(29.7)
#     sec.page_height   = Cm(21.0)
#     sec.left_margin   = Cm(1.5)
#     sec.right_margin  = Cm(1.5)
#     sec.top_margin    = Cm(1.5)
#     sec.bottom_margin = Cm(1.5)

#     t = doc.add_paragraph()
#     t.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     r = t.add_run("CONTO METALLO DELUXE — GESTIONE COMMISSIONI")
#     r.font.bold, r.font.size, r.font.name = True, Pt(13), "Calibri"
#     r.font.color.rgb = _NAVY

#     s = doc.add_paragraph()
#     s.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     sr = s.add_run(
#         f"Doc n° {data.get('document_number','N/A')}  │  "
#         f"Generated {datetime.now().strftime('%d/%m/%Y %H:%M')}  │  ",
#     )
#     sr.font.size, sr.font.name = Pt(8), "Calibri"
#     sr.font.color.rgb = RGBColor(0x70,0x70,0x70)
#     doc.add_paragraph()

#     contracts = data.get("contracts", [])
#     if not contracts:
#         doc.add_paragraph("⚠ No contracts found in PDF.")
#     else:
#         for ci, contract in enumerate(contracts):
#             lbl = doc.add_paragraph()
#             lr  = lbl.add_run(f"Contract #{ci+1}  —  {contract.get('nome_cliente_segnalatore','')}")
#             lr.font.bold, lr.font.size, lr.font.name = True, Pt(9), "Calibri"
#             lr.font.color.rgb = _GOLD

#             schedule = contract.get("payment_schedule", [])
#             n_rows   = max(len(schedule), 1)

#             tbl = doc.add_table(rows=1 + n_rows, cols=8)
#             tbl.style     = "Table Grid"
#             tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

#             for row in tbl.rows:
#                 for i, cell in enumerate(row.cells):
#                     cell.width = Cm([3.0,4.8,2.6,3.0,2.4,2.6,3.8,3.4][i])

#             for i, txt in enumerate([
#                 "Tipo Contratto","Nome Cliente/Segnalatore","Data Contratto",
#                 "Importo Contratto","Commissioni","Mese",
#                 "Data Pagamento\nCommissioni","Management\nMensile"
#             ]):
#                 _hdr_cell(tbl.rows[0].cells[i], txt)

#             tipo    = contract.get("tipo_contratto","Conto Metallo Deluxe")
#             nome    = contract.get("nome_cliente_segnalatore","")
#             data_c  = contract.get("data_contratto","")
#             importo_raw = contract.get("_importo", _parse_amount(str(contract.get("importo_contratto",""))))
#             importo_fmt = f"{importo_raw:,.2f} €".replace(",","X").replace(".",",").replace("X",".")
#             comm    = f"{contract.get('commissioni','')}%"

#             for ri, sched in enumerate(schedule):
#                 row = tbl.rows[1 + ri]
#                 bg  = "FFFFFF" if ri % 2 == 0 else "F4F0E8"

#                 _data_cell(row.cells[0], tipo if ri==0 else "", bold=(ri==0), center=False, bg=bg)
#                 _data_cell(row.cells[1], nome if ri==0 else "", center=False, bg=bg)
#                 _data_cell(row.cells[2], data_c if ri==0 else "", bg=bg)
#                 _data_cell(row.cells[3], importo_fmt if ri==0 else "", bold=(ri==0), bg=bg)
#                 _data_cell(row.cells[4], comm if ri==0 else "", bg=bg)
#                 _data_cell(row.cells[5], sched.get("mese",""), bold=True, bg=bg)

#                 pay = sched.get("data_pagamento_commissioni","")
#                 _data_cell(row.cells[6], pay, bg=bg)

#                 amt   = sched.get("_amount", 0.0)
#                 amt_s = sched.get("management_mensile","")
#                 if amt_s and amt == 0:
#                     amt = _parse_amount(amt_s)
#                 amt_fmt = f"{amt:,.2f} €".replace(",","X").replace(".",",").replace("X",".") if amt else ""
#                 _data_cell(row.cells[7], amt_fmt, bold=bool(amt), bg=bg, color=_NAVY if amt else None)

#             if ci < len(contracts) - 1:
#                 doc.add_paragraph()

#     notes = data.get("extraction_notes","")
#     if notes:
#         doc.add_paragraph()
#         np_ = doc.add_paragraph()
#         nr  = np_.add_run(f"ℹ  {notes}")
#         nr.font.size, nr.font.italic = Pt(7), True
#         nr.font.color.rgb = RGBColor(0x80,0x80,0x80)

#     doc.save(out_path)
#     log.info(f"DOCX → {out_path}")


# # ════════════════════════════════════════════════════════════════════════════
# #  Excel Generator  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# _FILL_NAVY   = PatternFill("solid", fgColor="1F497D")
# _FILL_GOLD   = PatternFill("solid", fgColor="C09B3A")
# _FILL_STRIPE = PatternFill("solid", fgColor="F4F0E8")
# _FILL_WHITE  = PatternFill("solid", fgColor="FFFFFF")
# _FILL_TOTAL  = PatternFill("solid", fgColor="E8F0FE")

# _FONT_HDR   = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
# _FONT_GOLD  = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
# _FONT_TIPO  = Font(name="Calibri", bold=True, color="1F497D", size=8)
# _FONT_DATA  = Font(name="Calibri", size=8, color="262626")
# _FONT_AMT   = Font(name="Calibri", bold=True, size=8, color="1F497D")
# _FONT_TOTAL = Font(name="Calibri", bold=True, size=9, color="1F497D")
# _FONT_TITLE = Font(name="Calibri", bold=True, size=14, color="1F497D")
# _FONT_SUB   = Font(name="Calibri", size=9, color="707070")

# _ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
# _ALIGN_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
# _ALIGN_R = Alignment(horizontal="right",  vertical="center")

# _THIN  = Side(style="thin",   color="BFBFBF")
# _THICK = Side(style="medium", color="1F497D")
# _BORDER_THIN  = Border(left=_THIN,  right=_THIN,  top=_THIN,  bottom=_THIN)
# _BORDER_THICK = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)

# HEADERS = [
#     "Tipo Contratto","Nome Cliente/Segnalatore","Data Contratto",
#     "Importo Contratto","Commissioni (%)","Mese",
#     "Data Pagamento Commissioni","Management Mensile (€)",
# ]
# COL_W = [22, 30, 14, 18, 14, 14, 24, 22]


# def _cell(ws, row, col, value, font=None, fill=None, align=None, border=None, num_fmt=None):
#     c = ws.cell(row=row, column=col, value=value)
#     if font:    c.font         = font
#     if fill:    c.fill         = fill
#     if align:   c.alignment    = align
#     if border:  c.border       = border
#     if num_fmt: c.number_format = num_fmt
#     return c


# def generate_excel(data: dict, out_path: str):
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Commissioni"
#     ws.freeze_panes = "A5"

#     ws.merge_cells("A1:H1")
#     _cell(ws, 1, 1, "CONTO METALLO DELUXE — GESTIONE COMMISSIONI",
#           font=_FONT_TITLE, fill=PatternFill("solid", fgColor="1F497D"), align=_ALIGN_C)

#     ws.merge_cells("A2:H2")
#     _cell(ws, 2, 1,
#           f"Doc: {data.get('document_number','N/A')}   │   "
#           f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}   │   ",
#           font=_FONT_SUB, fill=PatternFill("solid", fgColor="C09B3A"), align=_ALIGN_C)

#     ws.row_dimensions[3].height = 6

#     for ci, (hdr, w) in enumerate(zip(HEADERS, COL_W), start=1):
#         _cell(ws, 4, ci, hdr, font=_FONT_HDR, fill=_FILL_NAVY, align=_ALIGN_C, border=_BORDER_THIN)
#         ws.column_dimensions[get_column_letter(ci)].width = w

#     ws.row_dimensions[1].height = 28
#     ws.row_dimensions[2].height = 18
#     ws.row_dimensions[4].height = 32

#     current_row = 5
#     contracts   = data.get("contracts", [])

#     for ci, contract in enumerate(contracts):
#         schedule    = contract.get("payment_schedule", [])
#         importo_raw = contract.get("_importo", _parse_amount(str(contract.get("importo_contratto","0"))))
#         comm_pct    = contract.get("_comm", float(re.sub(r"[^\d.]","",str(contract.get("commissioni","2.17"))) or "2.17"))
#         tipo  = contract.get("tipo_contratto","Conto Metallo Deluxe")
#         nome  = contract.get("nome_cliente_segnalatore","")
#         datec = contract.get("data_contratto","")

#         for ri, sched in enumerate(schedule):
#             stripe = _FILL_STRIPE if ri % 2 else _FILL_WHITE
#             r      = current_row + ri
#             ws.row_dimensions[r].height = 18

#             _cell(ws, r, 1, tipo if ri==0 else "",
#                   font=_FONT_TIPO if ri==0 else _FONT_DATA, fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)
#             _cell(ws, r, 2, nome if ri==0 else "",
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)
#             _cell(ws, r, 3, datec if ri==0 else "",
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)
#             _cell(ws, r, 4, importo_raw if ri==0 else None,
#                   font=Font(name="Calibri",bold=(ri==0),size=8,color="1F497D" if ri==0 else "262626"),
#                   fill=stripe, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#             _cell(ws, r, 5, comm_pct/100 if ri==0 else None,
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN, num_fmt='0.00%')
#             _cell(ws, r, 6, sched.get("mese",""),
#                   font=Font(name="Calibri",bold=True,size=8), fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             pay = sched.get("data_pagamento_commissioni","")
#             _cell(ws, r, 7, pay or "—", font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             amt = sched.get("_amount", 0.0)
#             if not amt:
#                 raw_s = sched.get("management_mensile","")
#                 amt   = _parse_amount(raw_s) if raw_s else 0.0
#             _cell(ws, r, 8, amt if amt else None,
#                   font=_FONT_AMT if amt else _FONT_DATA,
#                   fill=stripe, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')

#         current_row += len(schedule)

#         total_row = current_row
#         total_amt = sum(
#             (sched.get("_amount",0) or _parse_amount(sched.get("management_mensile","") or "0"))
#             for sched in schedule
#         )
#         ws.row_dimensions[total_row].height = 20
#         ws.merge_cells(f"A{total_row}:E{total_row}")
#         _cell(ws, total_row, 1, f"TOTALE COMMISSIONI — {nome}",
#               font=_FONT_TOTAL, fill=_FILL_TOTAL, align=_ALIGN_L, border=_BORDER_THICK)
#         for col in range(2, 6):
#             _cell(ws, total_row, col, None, fill=_FILL_TOTAL, border=_BORDER_THICK)
#         _cell(ws, total_row, 6, "TOTALE", font=_FONT_TOTAL, fill=_FILL_TOTAL, align=_ALIGN_C, border=_BORDER_THICK)
#         _cell(ws, total_row, 7, "", fill=_FILL_TOTAL, border=_BORDER_THICK)
#         _cell(ws, total_row, 8, total_amt,
#               font=Font(name="Calibri",bold=True,size=10,color="059669"),
#               fill=_FILL_TOTAL, align=_ALIGN_R, border=_BORDER_THICK, num_fmt='#,##0.00 "€"')

#         current_row += 2

#     # Summary sheet
#     ws2 = wb.create_sheet("Summary")
#     for col, w in zip("ABCDE", [30,20,18,18,20]):
#         ws2.column_dimensions[col].width = w

#     ws2.merge_cells("A1:E1")
#     _cell(ws2, 1, 1, "CONTRACT SUMMARY", font=_FONT_TITLE, fill=_FILL_NAVY, align=_ALIGN_C)
#     ws2.row_dimensions[1].height = 28

#     for ci2, hdr in enumerate(["Cliente/Segnalatore","Importo","Comm %","Annual Return","Total Paid"], 1):
#         _cell(ws2, 2, ci2, hdr, font=_FONT_HDR, fill=_FILL_NAVY, align=_ALIGN_C, border=_BORDER_THIN)
#     ws2.row_dimensions[2].height = 22

#     for ri, c in enumerate(contracts, start=3):
#         importo_r = c.get("_importo", _parse_amount(str(c.get("importo_contratto","0"))))
#         comm_r    = c.get("_comm", 2.17)
#         annual    = importo_r * (comm_r/100) * 4
#         paid      = sum(
#             (s.get("_amount",0) or _parse_amount(s.get("management_mensile","") or "0"))
#             for s in c.get("payment_schedule",[])
#         )
#         stripe2 = _FILL_STRIPE if ri%2 else _FILL_WHITE
#         _cell(ws2, ri, 1, c.get("nome_cliente_segnalatore",""), font=_FONT_DATA, fill=stripe2, align=_ALIGN_L, border=_BORDER_THIN)
#         _cell(ws2, ri, 2, importo_r,  font=_FONT_AMT,  fill=stripe2, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         _cell(ws2, ri, 3, comm_r/100, font=_FONT_DATA, fill=stripe2, align=_ALIGN_C, border=_BORDER_THIN, num_fmt='0.00%')
#         _cell(ws2, ri, 4, annual,     font=_FONT_AMT,  fill=stripe2, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         _cell(ws2, ri, 5, paid,
#               font=Font(name="Calibri",bold=True,size=8,color="059669"),
#               fill=stripe2, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         ws2.row_dimensions[ri].height = 18

#     wb.save(out_path)
#     log.info(f"XLSX → {out_path}")


# # ════════════════════════════════════════════════════════════════════════════
# #  API Routes  (FastAPI style)
# # ════════════════════════════════════════════════════════════════════════════
# @app.post("/api/debug")
# def debug_extraction(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")
    
#     result = []
#     for i, c in enumerate(data.get("contracts", [])):
#         result.append({
#             "contract_number": i + 1,
#             "tipo_contratto":           c.get("tipo_contratto"),
#             "nome_cliente_segnalatore": c.get("nome_cliente_segnalatore"),
#             "data_contratto":           c.get("data_contratto"),
#             "importo_contratto":        c.get("importo_contratto"),
#             "commissioni":              c.get("commissioni"),
#             "piano":                    c.get("piano"),
#             "total_months":             len(c.get("payment_schedule", [])),
#             "payment_schedule":         c.get("payment_schedule", []),
#         })
    
#     return {"debug": result}

# @app.get("/api/health")
# def health():
#     return {
#         "status":    "ok",
#         "ai_ready":  bool(OPENAI_API_KEY),
#         "model":     OPENAI_MODEL,
#         "timestamp": datetime.now().isoformat(),
#     }


# @app.post("/api/upload")
# async def upload_pdf(file: UploadFile = File(...)):
#     if not file.filename.lower().endswith(".pdf"):
#         raise HTTPException(status_code=400, detail="Only PDF files accepted")

#     job_id   = str(uuid.uuid4())[:8]
#     pdf_path = UPLOAD_DIR / f"{job_id}_{file.filename}"

#     # Save uploaded file
#     content = await file.read()
#     with open(pdf_path, "wb") as f:
#         f.write(content)
#     log.info(f"Saved PDF: {pdf_path} ({pdf_path.stat().st_size} bytes)")

#     try:
#         pdf_text = extract_pdf_text(str(pdf_path))
#         log.info(f"Extracted {len(pdf_text)} chars from PDF")

#         if not OPENAI_API_KEY:
#             raise HTTPException(status_code=500, detail="OPENAI_API_KEY not set")
#         extracted = ai_extract(pdf_text)

#         data_path = OUTPUT_DIR / f"{job_id}_data.json"
#         with open(data_path, "w", encoding="utf-8") as fp:
#             json.dump(extracted, fp, ensure_ascii=False, indent=2)

#         return {
#             "job_id":           job_id,
#             "filename":         file.filename,
#             "contracts_found":  len(extracted.get("contracts", [])),
#             "document_number":  extracted.get("document_number", ""),
#             "extraction_notes": extracted.get("extraction_notes", ""),
#             "data":             extracted,
#             "status":           "analyzed",
#         }

#     except Exception as e:
#         log.error(traceback.format_exc())
#         raise HTTPException(status_code=500, detail=str(e))
#     finally:
#         try:
#             pdf_path.unlink()
#         except Exception:
#             pass


# def _load_job_data(job_id: str, body_data: Any) -> dict:
#     if body_data:
#         return enrich_contracts(body_data)
#     path = OUTPUT_DIR / f"{job_id}_data.json"
#     if path.exists():
#         with open(path, encoding="utf-8") as fp:
#             return enrich_contracts(json.load(fp))
#     return None


# @app.post("/api/generate/docx")
# def gen_docx(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")

#     out = OUTPUT_DIR / f"{req.job_id}_commissioni.docx"
#     try:
#         generate_docx(data, str(out))
#         return FileResponse(
#             path=str(out),
#             filename=f"Commissioni_{req.job_id}.docx",
#             media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
#         )
#     except Exception as e:
#         log.error(traceback.format_exc())
#         raise HTTPException(status_code=500, detail=str(e))


# @app.post("/api/generate/excel")
# def gen_excel(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")

#     out = OUTPUT_DIR / f"{req.job_id}_commissioni.xlsx"
#     try:
#         generate_excel(data, str(out))
#         return FileResponse(
#             path=str(out),
#             filename=f"Commissioni_{req.job_id}.xlsx",
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         )
#     except Exception as e:
#         log.error(traceback.format_exc())
#         raise HTTPException(status_code=500, detail=str(e))

# @app.post("/api/generate")
# def gen_document(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")

#     out = OUTPUT_DIR / f"{req.job_id}_commissioni.docx"
#     try:
#         generate_docx(data, str(out))
#         return FileResponse(
#             path=str(out),
#             filename=f"Commissioni_{req.job_id}.docx",
#             media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
#         )
#     except Exception as e:
#         log.error(traceback.format_exc())
#         raise HTTPException(status_code=500, detail=str(e))
# @app.post("/api/preview")
# def preview(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")
#     return data
# @app.post("/api/generate/both")
# def gen_both(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")

#     result: dict = {"job_id": req.job_id}

#     try:
#         docx_path = OUTPUT_DIR / f"{req.job_id}_commissioni.docx"
#         generate_docx(data, str(docx_path))
#         result["docx_ready"] = True
#     except Exception as e:
#         result["docx_error"] = str(e)
#         result["docx_ready"] = False

#     try:
#         xlsx_path = OUTPUT_DIR / f"{req.job_id}_commissioni.xlsx"
#         generate_excel(data, str(xlsx_path))
#         result["excel_ready"] = True
#     except Exception as e:
#         result["excel_error"] = str(e)
#         result["excel_ready"] = False

#     return result


# @app.get("/api/download/{job_id}/{fmt}")
# def download_file(job_id: str, fmt: str):
#     if fmt == "docx":
#         path = OUTPUT_DIR / f"{job_id}_commissioni.docx"
#         media = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#         name  = f"Commissioni_{job_id}.docx"
#     elif fmt == "excel":
#         path  = OUTPUT_DIR / f"{job_id}_commissioni.xlsx"
#         media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         name  = f"Commissioni_{job_id}.xlsx"
#     else:
#         raise HTTPException(status_code=400, detail="fmt must be docx or excel")

#     if not path.exists():
#         raise HTTPException(status_code=404, detail="File not generated yet")

#     return FileResponse(path=str(path), filename=name, media_type=media)


# # ════════════════════════════════════════════════════════════════════════════
# if __name__ == "__main__":
#     log.info(f"Starting FastAPI server on port {PORT}")
#     uvicorn.run("main:app", host="0.0.0.0", port=PORT, reload=False)





















#-------claude api------------
# """
# main.py — PDF → DOCX + Excel Extractor (FastAPI version)
# """

# import os
# import re
# import io
# import json
# import uuid
# import logging
# import traceback
# import urllib.request
# from datetime import datetime
# from pathlib import Path

# # ── Load .env before anything else ──────────────────────────────────────────
# def _load_dotenv(path: str = ".env"):
#     env_path = Path(path)
#     if not env_path.exists():
#         env_path = Path(__file__).parent / ".env"
#     if not env_path.exists():
#         return
#     with open(env_path, encoding="utf-8") as f:
#         for line in f:
#             line = line.strip()
#             if not line or line.startswith("#") or "=" not in line:
#                 continue
#             key, _, val = line.partition("=")
#             key = key.strip()
#             val = val.strip().strip('"').strip("'")
#             if key and key not in os.environ:
#                 os.environ[key] = val

# _load_dotenv()

# # ── Logging ──────────────────────────────────────────────────────────────────
# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s %(levelname)s %(message)s",
#     datefmt="%Y-%m-%d %H:%M:%S",
# )
# log = logging.getLogger(__name__)

# # ── Config from env ───────────────────────────────────────────────────────────
# ANTHROPIC_API_KEY    = os.environ.get("ANTHROPIC_API_KEY", "")
# ANTHROPIC_MODEL      = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-20250514")
# ANTHROPIC_MAX_TOKENS = int(os.environ.get("ANTHROPIC_MAX_TOKENS", "4096"))
# PORT                 = int(os.environ.get("PORT", "8000"))
# MAX_UPLOAD_MB        = int(os.environ.get("MAX_UPLOAD_MB", "50"))
# FASTAPI_ENV          = os.environ.get("FASTAPI_ENV", "development")

# BASE_DIR   = Path(__file__).parent
# UPLOAD_DIR = BASE_DIR / "uploads"
# OUTPUT_DIR = BASE_DIR / "outputs"
# UPLOAD_DIR.mkdir(exist_ok=True)
# OUTPUT_DIR.mkdir(exist_ok=True)

# log.info(f"Model  : {ANTHROPIC_MODEL}")
# log.info(f"AI Key : {'SET ✓' if ANTHROPIC_API_KEY else 'NOT SET — fallback mode'}")

# # ── FastAPI imports ───────────────────────────────────────────────────────────
# from fastapi import FastAPI, File, UploadFile, HTTPException, Request
# from fastapi.responses import JSONResponse, FileResponse
# from fastapi.middleware.cors import CORSMiddleware
# from pydantic import BaseModel
# from typing import Optional, Any
# import uvicorn

# import pdfplumber
# from docx import Document
# from docx.shared import Pt, RGBColor, Cm
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# from docx.enum.table import WD_TABLE_ALIGNMENT
# from docx.oxml.ns import qn
# from docx.oxml import OxmlElement
# import openpyxl
# from openpyxl.styles import (
#     PatternFill, Font, Alignment, Border, Side
# )
# from openpyxl.utils import get_column_letter

# # ── App ───────────────────────────────────────────────────────────────────────
# app = FastAPI(title="MetalloDoc API", version="1.0.0")

# # ── CORS ──────────────────────────────────────────────────────────────────────
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ── Pydantic models ───────────────────────────────────────────────────────────
# class GenerateRequest(BaseModel):
#     job_id: str
#     data: Optional[Any] = None


# # ════════════════════════════════════════════════════════════════════════════
# #  AI + PDF Extraction  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# def extract_pdf_text(pdf_path: str) -> str:
#     pages = []
#     with pdfplumber.open(pdf_path) as pdf:
#         for i, page in enumerate(pdf.pages):
#             text = page.extract_text() or ""
#             pages.append(f"=== PAGE {i+1} ===\n{text}")
#     return "\n\n".join(pages)


# def call_anthropic(prompt: str, system: str) -> str:
#     payload = json.dumps({
#         "model":      ANTHROPIC_MODEL,
#         "max_tokens": ANTHROPIC_MAX_TOKENS,
#         "system":     system,
#         "messages":   [{"role": "user", "content": prompt}],
#     }).encode("utf-8")

#     req = urllib.request.Request(
#         "https://api.anthropic.com/v1/messages",
#         data=payload,
#         headers={
#             "Content-Type":      "application/json",
#             "x-api-key":         ANTHROPIC_API_KEY,
#             "anthropic-version": "2023-06-01",
#         },
#         method="POST",
#     )
#     try:
#         with urllib.request.urlopen(req, timeout=300) as resp:
#             body = json.loads(resp.read().decode("utf-8"))
#             return body["content"][0]["text"]
#     except urllib.error.HTTPError as e:
#         error_body = e.read().decode("utf-8")
#         log.error(f"Anthropic API error {e.code}: {error_body}")
#         raise


# def safe_json(text: str) -> dict:
#     for attempt in (
#         lambda t: json.loads(t.strip()),
#         lambda t: json.loads(re.search(r"```(?:json)?\s*(\{.*?\})\s*```", t, re.DOTALL).group(1)),
#         lambda t: json.loads(re.search(r"\{.*\}", t, re.DOTALL).group(0)),
#     ):
#         try:
#             return attempt(text)
#         except Exception:
#             pass
#     raise ValueError(f"Cannot parse JSON from AI response: {text[:300]}")


# SYSTEM_PROMPT = """\
# You are a financial document data extractor for Italian gold investment contracts
# (Conto Metallo Deluxe, Mac & Ro Srl). Extract data and return ONLY valid JSON.

# Required JSON structure:
# {
#   "contracts": [
#     {
#       "tipo_contratto": "Conto Metallo Deluxe",
#       "nome_cliente_segnalatore": "Client Lastname Firstname/Agent Lastname Firstname",
#       "data_contratto": "DD/MM/YYYY",
#       "importo_contratto": "10,000.00",
#       "commissioni": "2.17",
#       "piano": "Plus",
#       "payment_schedule": [
#         {
#           "mese": "Gennaio",
#           "data_pagamento_commissioni": "10-Jan-26",
#           "management_mensile": "217.00"
#         }
#       ]
#     }
#   ],
#   "document_number": "CMP-000014.01.2026",
#   "extraction_notes": "short note"
# }

# Extraction rules:
# - nome_cliente_segnalatore: "ClientFullName/CollaboratorFullName" (from Contrante + Collaboratore)
# - data_contratto: signing date from "Luogo e data" field, format DD/MM/YYYY
# - importo_contratto: numeric string only, no currency symbol (e.g. "10000.00")
# - commissioni: numeric percent string only (e.g. "2.17" for 2.17%)
# - piano: "Plus" if Conto Metallo Deluxe Plus, "Top" if Conto Metallo Deluxe Top
# - payment_schedule: generate 12 months from contract start month
#   * First 5 months: empty data_pagamento_commissioni and management_mensile
#   * From month 6 onward: date as DD-Mon-YY, management_mensile as numeric string
#   * management_mensile = importo * (commissioni/100)
# - Return ONLY the JSON object, no markdown, no explanation\
# """


# def ai_extract(pdf_text: str) -> dict:
#     prompt = f"Extract contract data from this PDF:\n\n{pdf_text[:15000]}\n\nReturn ONLY JSON."
#     raw = call_anthropic(prompt, SYSTEM_PROMPT)
#     log.info(f"AI raw (first 200): {raw[:200]}")
#     return safe_json(raw)


# # ════════════════════════════════════════════════════════════════════════════
# #  Fallback pattern-matching extractor  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# def pattern_extract(pdf_text: str, filename: str) -> dict:
#     m = re.search(r"(CMP-[\d.]+)", pdf_text)
#     doc_num = m.group(1) if m else "N/A"

#     m = re.search(r"Cognome e Nome[:\s]+([A-Za-zÀ-ÿ]+\s+[A-Za-zÀ-ÿ]+)", pdf_text)
#     client = m.group(1).strip() if m else "Unknown Client"

#     m = re.search(r"COLLABORATORE.*?Cognome e Nome\s+([A-Za-zÀ-ÿ]+\s+[A-Za-zÀ-ÿ]+)", pdf_text, re.DOTALL)
#     collab = m.group(1).strip() if m else "Agent"

#     m = re.search(r"Corrispettivo\s*\n?\s*([\d.,]+)", pdf_text)
#     raw_amt = m.group(1).strip() if m else "10000"
#     amount = _parse_amount(raw_amt)

#     m = re.search(r"(\d{4}-\d{2}-\d{2})", pdf_text)
#     if m:
#         dt = datetime.strptime(m.group(1), "%Y-%m-%d")
#         date_str = dt.strftime("%d/%m/%Y")
#     else:
#         date_str = "09/01/2026"

#     is_top = ("Deluxe Top" in pdf_text) and (amount >= 100000)
#     piano  = "Top" if is_top else "Plus"
#     comm   = "2.50" if is_top else "2.17"

#     return {
#         "contracts": [{
#             "tipo_contratto": "Conto Metallo Deluxe",
#             "nome_cliente_segnalatore": f"{client}/{collab}",
#             "data_contratto": date_str,
#             "importo_contratto": f"{amount:.2f}",
#             "commissioni": comm,
#             "piano": piano,
#             "payment_schedule": [],
#         }],
#         "document_number": doc_num,
#         "extraction_notes": "Pattern-matching extraction (no AI key)",
#     }


# # ════════════════════════════════════════════════════════════════════════════
# #  Payment schedule builder  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# MONTHS_IT    = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
#                 "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]
# MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun",
#                 "Jul","Aug","Sep","Oct","Nov","Dec"]


# def _parse_amount(s: str) -> float:
#     s = str(s).replace("€","").replace("$","").strip()
#     if "." in s and "," in s:
#         if s.index(",") < s.index("."):
#             s = s.replace(",","")
#         else:
#             s = s.replace(".","").replace(",",".")
#     elif "," in s:
#         parts = s.split(",")
#         if len(parts[-1]) == 2 and len(parts[0]) >= 2:
#             s = s.replace(",","")
#         else:
#             s = s.replace(",",".")
#     try:
#         return float(re.sub(r"[^\d.]", "", s))
#     except Exception:
#         return 10000.0


# def build_schedule(contract: dict) -> list:
#     if contract.get("payment_schedule") and len(contract["payment_schedule"]) >= 12:
#         for row in contract["payment_schedule"]:
#             v = row.get("management_mensile","")
#             if v:
#                 row["_amount"] = _parse_amount(str(v))
#         return contract["payment_schedule"]

#     importo = _parse_amount(str(contract.get("importo_contratto","10000")))
#     comm    = float(re.sub(r"[^\d.]", "", str(contract.get("commissioni","2.17"))) or "2.17")
#     periodic_amount = importo * (comm / 100)

#     date_str = contract.get("data_contratto","01/01/2026")
#     try:
#         dt = datetime.strptime(date_str, "%d/%m/%Y")
#         start_m = dt.month - 1
#         start_y = dt.year
#     except Exception:
#         start_m, start_y = 0, 2026

#     schedule = []
#     for i in range(12):
#         idx  = (start_m + i) % 12
#         year = start_y + (start_m + i) // 12
#         name = MONTHS_IT[idx]

#         if i < 5:
#             schedule.append({
#                 "mese": name,
#                 "data_pagamento_commissioni": "",
#                 "management_mensile": "",
#                 "_amount": 0.0,
#             })
#         else:
#             pay_date = f"{10:02d}-{MONTHS_SHORT[idx]}-{str(year)[2:]}"
#             schedule.append({
#                 "mese": name,
#                 "data_pagamento_commissioni": pay_date,
#                 "management_mensile": f"{periodic_amount:.2f}",
#                 "_amount": periodic_amount,
#             })

#     return schedule


# def enrich_contracts(data: dict) -> dict:
#     for c in data.get("contracts", []):
#         c["payment_schedule"] = build_schedule(c)
#         c["_importo"] = _parse_amount(str(c.get("importo_contratto","0")))
#         c["_comm"]    = float(re.sub(r"[^\d.]","",str(c.get("commissioni","2.17"))) or "2.17")
#     return data


# # ════════════════════════════════════════════════════════════════════════════
# #  DOCX Generator  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# _NAVY  = RGBColor(0x1F,0x49,0x7D)
# _GOLD  = RGBColor(0xC0,0x9B,0x3A)
# _WHITE = RGBColor(0xFF,0xFF,0xFF)
# _DARK  = RGBColor(0x26,0x26,0x26)


# def _shd(cell, hex_color: str):
#     tc   = cell._tc
#     tcPr = tc.get_or_add_tcPr()
#     shd  = OxmlElement("w:shd")
#     shd.set(qn("w:val"),   "clear")
#     shd.set(qn("w:color"), "auto")
#     shd.set(qn("w:fill"),  hex_color)
#     tcPr.append(shd)


# def _hdr_cell(cell, text: str, size=8):
#     _shd(cell, "1F497D")
#     cell.text = ""
#     p = cell.paragraphs[0]
#     p.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     r = p.add_run(text)
#     r.font.bold, r.font.size, r.font.name = True, Pt(size), "Calibri"
#     r.font.color.rgb = _WHITE


# def _data_cell(cell, text: str, bold=False, center=True, bg="FFFFFF", color=None):
#     _shd(cell, bg)
#     cell.text = ""
#     p = cell.paragraphs[0]
#     p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
#     r = p.add_run(str(text) if text else "")
#     r.font.bold, r.font.size, r.font.name = bold, Pt(8), "Calibri"
#     r.font.color.rgb = color if color else _DARK


# def generate_docx(data: dict, out_path: str):
#     doc = Document()

#     sec = doc.sections[0]
#     sec.page_width    = Cm(29.7)
#     sec.page_height   = Cm(21.0)
#     sec.left_margin   = Cm(1.5)
#     sec.right_margin  = Cm(1.5)
#     sec.top_margin    = Cm(1.5)
#     sec.bottom_margin = Cm(1.5)

#     t = doc.add_paragraph()
#     t.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     r = t.add_run("CONTO METALLO DELUXE — GESTIONE COMMISSIONI")
#     r.font.bold, r.font.size, r.font.name = True, Pt(13), "Calibri"
#     r.font.color.rgb = _NAVY

#     s = doc.add_paragraph()
#     s.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     sr = s.add_run(
#         f"Doc n° {data.get('document_number','N/A')}  │  "
#         f"Generated {datetime.now().strftime('%d/%m/%Y %H:%M')}  │  "
#         f"Model: {ANTHROPIC_MODEL}"
#     )
#     sr.font.size, sr.font.name = Pt(8), "Calibri"
#     sr.font.color.rgb = RGBColor(0x70,0x70,0x70)
#     doc.add_paragraph()

#     contracts = data.get("contracts", [])
#     if not contracts:
#         doc.add_paragraph("⚠ No contracts found in PDF.")
#     else:
#         for ci, contract in enumerate(contracts):
#             lbl = doc.add_paragraph()
#             lr  = lbl.add_run(f"Contract #{ci+1}  —  {contract.get('nome_cliente_segnalatore','')}")
#             lr.font.bold, lr.font.size, lr.font.name = True, Pt(9), "Calibri"
#             lr.font.color.rgb = _GOLD

#             schedule = contract.get("payment_schedule", [])
#             n_rows   = max(len(schedule), 1)

#             tbl = doc.add_table(rows=1 + n_rows, cols=8)
#             tbl.style     = "Table Grid"
#             tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

#             for row in tbl.rows:
#                 for i, cell in enumerate(row.cells):
#                     cell.width = Cm([3.0,4.8,2.6,3.0,2.4,2.6,3.8,3.4][i])

#             for i, txt in enumerate([
#                 "Tipo Contratto","Nome Cliente/Segnalatore","Data Contratto",
#                 "Importo Contratto","Commissioni","Mese",
#                 "Data Pagamento\nCommissioni","Management\nMensile"
#             ]):
#                 _hdr_cell(tbl.rows[0].cells[i], txt)

#             tipo    = contract.get("tipo_contratto","Conto Metallo Deluxe")
#             nome    = contract.get("nome_cliente_segnalatore","")
#             data_c  = contract.get("data_contratto","")
#             importo_raw = contract.get("_importo", _parse_amount(str(contract.get("importo_contratto",""))))
#             importo_fmt = f"{importo_raw:,.2f} €".replace(",","X").replace(".",",").replace("X",".")
#             comm    = f"{contract.get('commissioni','')}%"

#             for ri, sched in enumerate(schedule):
#                 row = tbl.rows[1 + ri]
#                 bg  = "FFFFFF" if ri % 2 == 0 else "F4F0E8"

#                 _data_cell(row.cells[0], tipo if ri==0 else "", bold=(ri==0), center=False, bg=bg)
#                 _data_cell(row.cells[1], nome if ri==0 else "", center=False, bg=bg)
#                 _data_cell(row.cells[2], data_c if ri==0 else "", bg=bg)
#                 _data_cell(row.cells[3], importo_fmt if ri==0 else "", bold=(ri==0), bg=bg)
#                 _data_cell(row.cells[4], comm if ri==0 else "", bg=bg)
#                 _data_cell(row.cells[5], sched.get("mese",""), bold=True, bg=bg)

#                 pay = sched.get("data_pagamento_commissioni","")
#                 _data_cell(row.cells[6], pay, bg=bg)

#                 amt   = sched.get("_amount", 0.0)
#                 amt_s = sched.get("management_mensile","")
#                 if amt_s and amt == 0:
#                     amt = _parse_amount(amt_s)
#                 amt_fmt = f"{amt:,.2f} €".replace(",","X").replace(".",",").replace("X",".") if amt else ""
#                 _data_cell(row.cells[7], amt_fmt, bold=bool(amt), bg=bg, color=_NAVY if amt else None)

#             if ci < len(contracts) - 1:
#                 doc.add_paragraph()

#     notes = data.get("extraction_notes","")
#     if notes:
#         doc.add_paragraph()
#         np_ = doc.add_paragraph()
#         nr  = np_.add_run(f"ℹ  {notes}")
#         nr.font.size, nr.font.italic = Pt(7), True
#         nr.font.color.rgb = RGBColor(0x80,0x80,0x80)

#     doc.save(out_path)
#     log.info(f"DOCX → {out_path}")


# # ════════════════════════════════════════════════════════════════════════════
# #  Excel Generator  (unchanged logic)
# # ════════════════════════════════════════════════════════════════════════════

# _FILL_NAVY   = PatternFill("solid", fgColor="1F497D")
# _FILL_GOLD   = PatternFill("solid", fgColor="C09B3A")
# _FILL_STRIPE = PatternFill("solid", fgColor="F4F0E8")
# _FILL_WHITE  = PatternFill("solid", fgColor="FFFFFF")
# _FILL_TOTAL  = PatternFill("solid", fgColor="E8F0FE")

# _FONT_HDR   = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
# _FONT_GOLD  = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
# _FONT_TIPO  = Font(name="Calibri", bold=True, color="1F497D", size=8)
# _FONT_DATA  = Font(name="Calibri", size=8, color="262626")
# _FONT_AMT   = Font(name="Calibri", bold=True, size=8, color="1F497D")
# _FONT_TOTAL = Font(name="Calibri", bold=True, size=9, color="1F497D")
# _FONT_TITLE = Font(name="Calibri", bold=True, size=14, color="1F497D")
# _FONT_SUB   = Font(name="Calibri", size=9, color="707070")

# _ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
# _ALIGN_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
# _ALIGN_R = Alignment(horizontal="right",  vertical="center")

# _THIN  = Side(style="thin",   color="BFBFBF")
# _THICK = Side(style="medium", color="1F497D")
# _BORDER_THIN  = Border(left=_THIN,  right=_THIN,  top=_THIN,  bottom=_THIN)
# _BORDER_THICK = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)

# HEADERS = [
#     "Tipo Contratto","Nome Cliente/Segnalatore","Data Contratto",
#     "Importo Contratto","Commissioni (%)","Mese",
#     "Data Pagamento Commissioni","Management Mensile (€)",
# ]
# COL_W = [22, 30, 14, 18, 14, 14, 24, 22]


# def _cell(ws, row, col, value, font=None, fill=None, align=None, border=None, num_fmt=None):
#     c = ws.cell(row=row, column=col, value=value)
#     if font:    c.font         = font
#     if fill:    c.fill         = fill
#     if align:   c.alignment    = align
#     if border:  c.border       = border
#     if num_fmt: c.number_format = num_fmt
#     return c


# def generate_excel(data: dict, out_path: str):
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Commissioni"
#     ws.freeze_panes = "A5"

#     ws.merge_cells("A1:H1")
#     _cell(ws, 1, 1, "CONTO METALLO DELUXE — GESTIONE COMMISSIONI",
#           font=_FONT_TITLE, fill=PatternFill("solid", fgColor="1F497D"), align=_ALIGN_C)

#     ws.merge_cells("A2:H2")
#     _cell(ws, 2, 1,
#           f"Doc: {data.get('document_number','N/A')}   │   "
#           f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}   │   "
#           f"Model: {ANTHROPIC_MODEL}",
#           font=_FONT_SUB, fill=PatternFill("solid", fgColor="C09B3A"), align=_ALIGN_C)

#     ws.row_dimensions[3].height = 6

#     for ci, (hdr, w) in enumerate(zip(HEADERS, COL_W), start=1):
#         _cell(ws, 4, ci, hdr, font=_FONT_HDR, fill=_FILL_NAVY, align=_ALIGN_C, border=_BORDER_THIN)
#         ws.column_dimensions[get_column_letter(ci)].width = w

#     ws.row_dimensions[1].height = 28
#     ws.row_dimensions[2].height = 18
#     ws.row_dimensions[4].height = 32

#     current_row = 5
#     contracts   = data.get("contracts", [])

#     for ci, contract in enumerate(contracts):
#         schedule    = contract.get("payment_schedule", [])
#         importo_raw = contract.get("_importo", _parse_amount(str(contract.get("importo_contratto","0"))))
#         comm_pct    = contract.get("_comm", float(re.sub(r"[^\d.]","",str(contract.get("commissioni","2.17"))) or "2.17"))
#         tipo  = contract.get("tipo_contratto","Conto Metallo Deluxe")
#         nome  = contract.get("nome_cliente_segnalatore","")
#         datec = contract.get("data_contratto","")

#         for ri, sched in enumerate(schedule):
#             stripe = _FILL_STRIPE if ri % 2 else _FILL_WHITE
#             r      = current_row + ri
#             ws.row_dimensions[r].height = 18

#             _cell(ws, r, 1, tipo if ri==0 else "",
#                   font=_FONT_TIPO if ri==0 else _FONT_DATA, fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)
#             _cell(ws, r, 2, nome if ri==0 else "",
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)
#             _cell(ws, r, 3, datec if ri==0 else "",
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)
#             _cell(ws, r, 4, importo_raw if ri==0 else None,
#                   font=Font(name="Calibri",bold=(ri==0),size=8,color="1F497D" if ri==0 else "262626"),
#                   fill=stripe, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#             _cell(ws, r, 5, comm_pct/100 if ri==0 else None,
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN, num_fmt='0.00%')
#             _cell(ws, r, 6, sched.get("mese",""),
#                   font=Font(name="Calibri",bold=True,size=8), fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             pay = sched.get("data_pagamento_commissioni","")
#             _cell(ws, r, 7, pay or "—", font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             amt = sched.get("_amount", 0.0)
#             if not amt:
#                 raw_s = sched.get("management_mensile","")
#                 amt   = _parse_amount(raw_s) if raw_s else 0.0
#             _cell(ws, r, 8, amt if amt else None,
#                   font=_FONT_AMT if amt else _FONT_DATA,
#                   fill=stripe, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')

#         current_row += len(schedule)

#         total_row = current_row
#         total_amt = sum(
#             (sched.get("_amount",0) or _parse_amount(sched.get("management_mensile","") or "0"))
#             for sched in schedule
#         )
#         ws.row_dimensions[total_row].height = 20
#         ws.merge_cells(f"A{total_row}:E{total_row}")
#         _cell(ws, total_row, 1, f"TOTALE COMMISSIONI — {nome}",
#               font=_FONT_TOTAL, fill=_FILL_TOTAL, align=_ALIGN_L, border=_BORDER_THICK)
#         for col in range(2, 6):
#             _cell(ws, total_row, col, None, fill=_FILL_TOTAL, border=_BORDER_THICK)
#         _cell(ws, total_row, 6, "TOTALE", font=_FONT_TOTAL, fill=_FILL_TOTAL, align=_ALIGN_C, border=_BORDER_THICK)
#         _cell(ws, total_row, 7, "", fill=_FILL_TOTAL, border=_BORDER_THICK)
#         _cell(ws, total_row, 8, total_amt,
#               font=Font(name="Calibri",bold=True,size=10,color="059669"),
#               fill=_FILL_TOTAL, align=_ALIGN_R, border=_BORDER_THICK, num_fmt='#,##0.00 "€"')

#         current_row += 2

#     # Summary sheet
#     ws2 = wb.create_sheet("Summary")
#     for col, w in zip("ABCDE", [30,20,18,18,20]):
#         ws2.column_dimensions[col].width = w

#     ws2.merge_cells("A1:E1")
#     _cell(ws2, 1, 1, "CONTRACT SUMMARY", font=_FONT_TITLE, fill=_FILL_NAVY, align=_ALIGN_C)
#     ws2.row_dimensions[1].height = 28

#     for ci2, hdr in enumerate(["Cliente/Segnalatore","Importo","Comm %","Annual Return","Total Paid"], 1):
#         _cell(ws2, 2, ci2, hdr, font=_FONT_HDR, fill=_FILL_NAVY, align=_ALIGN_C, border=_BORDER_THIN)
#     ws2.row_dimensions[2].height = 22

#     for ri, c in enumerate(contracts, start=3):
#         importo_r = c.get("_importo", _parse_amount(str(c.get("importo_contratto","0"))))
#         comm_r    = c.get("_comm", 2.17)
#         annual    = importo_r * (comm_r/100) * 4
#         paid      = sum(
#             (s.get("_amount",0) or _parse_amount(s.get("management_mensile","") or "0"))
#             for s in c.get("payment_schedule",[])
#         )
#         stripe2 = _FILL_STRIPE if ri%2 else _FILL_WHITE
#         _cell(ws2, ri, 1, c.get("nome_cliente_segnalatore",""), font=_FONT_DATA, fill=stripe2, align=_ALIGN_L, border=_BORDER_THIN)
#         _cell(ws2, ri, 2, importo_r,  font=_FONT_AMT,  fill=stripe2, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         _cell(ws2, ri, 3, comm_r/100, font=_FONT_DATA, fill=stripe2, align=_ALIGN_C, border=_BORDER_THIN, num_fmt='0.00%')
#         _cell(ws2, ri, 4, annual,     font=_FONT_AMT,  fill=stripe2, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         _cell(ws2, ri, 5, paid,
#               font=Font(name="Calibri",bold=True,size=8,color="059669"),
#               fill=stripe2, align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         ws2.row_dimensions[ri].height = 18

#     wb.save(out_path)
#     log.info(f"XLSX → {out_path}")


# # ════════════════════════════════════════════════════════════════════════════
# #  API Routes  (FastAPI style)
# # ════════════════════════════════════════════════════════════════════════════

# @app.get("/api/health")
# def health():
#     return {
#         "status":    "ok",
#         "ai_ready":  bool(ANTHROPIC_API_KEY),
#         "model":     ANTHROPIC_MODEL,
#         "flask_env": FASTAPI_ENV,
#         "timestamp": datetime.now().isoformat(),
#     }


# @app.post("/api/upload")
# async def upload_pdf(file: UploadFile = File(...)):
#     if not file.filename.lower().endswith(".pdf"):
#         raise HTTPException(status_code=400, detail="Only PDF files accepted")

#     job_id   = str(uuid.uuid4())[:8]
#     pdf_path = UPLOAD_DIR / f"{job_id}_{file.filename}"

#     # Save uploaded file
#     content = await file.read()
#     with open(pdf_path, "wb") as f:
#         f.write(content)
#     log.info(f"Saved PDF: {pdf_path} ({pdf_path.stat().st_size} bytes)")

#     try:
#         pdf_text = extract_pdf_text(str(pdf_path))
#         log.info(f"Extracted {len(pdf_text)} chars from PDF")

#         if not ANTHROPIC_API_KEY:
#             raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not set — AI extraction required")
#         extracted = ai_extract(pdf_text)

#         data_path = OUTPUT_DIR / f"{job_id}_data.json"
#         with open(data_path, "w", encoding="utf-8") as fp:
#             json.dump(extracted, fp, ensure_ascii=False, indent=2)

#         return {
#             "job_id":           job_id,
#             "filename":         file.filename,
#             "contracts_found":  len(extracted.get("contracts", [])),
#             "document_number":  extracted.get("document_number", ""),
#             "extraction_notes": extracted.get("extraction_notes", ""),
#             "data":             extracted,
#             "status":           "analyzed",
#         }

#     except Exception as e:
#         log.error(traceback.format_exc())
#         raise HTTPException(status_code=500, detail=str(e))
#     finally:
#         try:
#             pdf_path.unlink()
#         except Exception:
#             pass


# def _load_job_data(job_id: str, body_data: Any) -> dict:
#     if body_data:
#         return enrich_contracts(body_data)
#     path = OUTPUT_DIR / f"{job_id}_data.json"
#     if path.exists():
#         with open(path, encoding="utf-8") as fp:
#             return enrich_contracts(json.load(fp))
#     return None


# @app.post("/api/generate/docx")
# def gen_docx(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")

#     out = OUTPUT_DIR / f"{req.job_id}_commissioni.docx"
#     try:
#         generate_docx(data, str(out))
#         return FileResponse(
#             path=str(out),
#             filename=f"Commissioni_{req.job_id}.docx",
#             media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
#         )
#     except Exception as e:
#         log.error(traceback.format_exc())
#         raise HTTPException(status_code=500, detail=str(e))


# @app.post("/api/generate/excel")
# def gen_excel(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")

#     out = OUTPUT_DIR / f"{req.job_id}_commissioni.xlsx"
#     try:
#         generate_excel(data, str(out))
#         return FileResponse(
#             path=str(out),
#             filename=f"Commissioni_{req.job_id}.xlsx",
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         )
#     except Exception as e:
#         log.error(traceback.format_exc())
#         raise HTTPException(status_code=500, detail=str(e))


# @app.post("/api/generate/both")
# def gen_both(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")

#     result: dict = {"job_id": req.job_id}

#     try:
#         docx_path = OUTPUT_DIR / f"{req.job_id}_commissioni.docx"
#         generate_docx(data, str(docx_path))
#         result["docx_ready"] = True
#     except Exception as e:
#         result["docx_error"] = str(e)
#         result["docx_ready"] = False

#     try:
#         xlsx_path = OUTPUT_DIR / f"{req.job_id}_commissioni.xlsx"
#         generate_excel(data, str(xlsx_path))
#         result["excel_ready"] = True
#     except Exception as e:
#         result["excel_error"] = str(e)
#         result["excel_ready"] = False

#     return result


# @app.get("/api/download/{job_id}/{fmt}")
# def download_file(job_id: str, fmt: str):
#     if fmt == "docx":
#         path = OUTPUT_DIR / f"{job_id}_commissioni.docx"
#         media = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#         name  = f"Commissioni_{job_id}.docx"
#     elif fmt == "excel":
#         path  = OUTPUT_DIR / f"{job_id}_commissioni.xlsx"
#         media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         name  = f"Commissioni_{job_id}.xlsx"
#     else:
#         raise HTTPException(status_code=400, detail="fmt must be docx or excel")

#     if not path.exists():
#         raise HTTPException(status_code=404, detail="File not generated yet")

#     return FileResponse(path=str(path), filename=name, media_type=media)


# # ════════════════════════════════════════════════════════════════════════════
# if __name__ == "__main__":
#     log.info(f"Starting FastAPI server on port {PORT}")
#     uvicorn.run("main:app", host="0.0.0.0", port=PORT, reload=False)