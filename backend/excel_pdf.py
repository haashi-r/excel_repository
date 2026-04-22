#----own self hosted ollama model--------------------
"""
main2.py — MetalloDoc API (Self-Hosted / Offline version)
LLM: Ollama (local, no internet, no cloud)
"""


import os
import re
import json
import uuid
import logging
import traceback
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

# ── Load .env ─────────────────────────────────────────────────────────────────
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
PORT          = int(os.environ.get("PORT", "8000"))
OLLAMA_HOST   = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
OLLAMA_MODEL  = os.environ.get("OLLAMA_MODEL", "qwen2.5:7b")
OLLAMA_TIMEOUT = int(os.environ.get("OLLAMA_TIMEOUT", "600"))

log.info(f"Ollama host  : {OLLAMA_HOST}")
log.info(f"Ollama model : {OLLAMA_MODEL}")

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ── FastAPI ───────────────────────────────────────────────────────────────────
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse
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
import calendar
from dateutil.relativedelta import relativedelta

app = FastAPI(title="MetalloDoc API (Offline)", version="3.0.0")
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
#  SYSTEM PROMPT
# ════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """\
You are a financial document data extractor for Italian gold investment contracts
(Conto Metallo Deluxe, Mac & Ro Srl). Extract data and return ONLY valid JSON.

SEARCH THE ENTIRE DOCUMENT CAREFULLY FOR THESE FIELDS:

1. importo_contratto
   - Search for "Corrispettivo" field
   - Also check "Importo previsto di investimento" field
   - Return numeric string only, no euro symbol
   - Example: "10000.00"

2. commissioni_annuale
   - Search Article 2 for percentage
   - Plus plan = "6.00", Top plan = "10.00"
   - Return numeric string only

3. frequenza_pagamento
   - Search Article 4 section 5
   - "semestrale" = every 6 months
   - "trimestrale" = every 3 months
   - "mensile" = every month
   - Return exact word: "semestrale" or "trimestrale" or "mensile"

4. percentuale_per_periodo
   - semestrale: commissioni_annuale / 2
   - trimestrale: commissioni_annuale / 4
   - mensile: commissioni_annuale / 12
   - Return numeric string only

5. data_contratto
   - Search "Luogo e data" field
   - Format: DD/MM/YYYY
   - Example: "2026-01-13" becomes "13/01/2026"

6. durata_mesi
   - Search "Durata" field
   - Return number only: "12" or "24" or "36" or "48" or "60"

7. piano
   - If "Conto Metallo Deluxe Plus" is checked return "Plus"
   - If "Conto Metallo Deluxe Top" is checked return "Top"

8. nome_cliente_segnalatore
   - Format: "CLIENT_FULLNAME/AGENT_FULLNAME"
   - Client = CONTRAENTE PERSONA FISICA Cognome e Nome
   - Agent = COLLABORATORE Cognome e Nome
   - Example: "QUARANTA PIERPAOLO/Plaitano Dario"

9. document_number
   - Search "Documento n" field
   - Example: "CMP-000018.01.2026"

IMPORTANT RULES:
- Search the ENTIRE document, not just first pages
- Extract ONLY values physically written in the PDF
- NEVER invent or guess any value
- If a field is not found, use empty string ""
- Return ONLY raw JSON, absolutely no markdown, no explanation, no code blocks

Return ONLY this exact JSON structure:
{
  "contracts": [
    {
      "tipo_contratto": "Conto Metallo Deluxe",
      "nome_cliente_segnalatore": "CLIENT/AGENT",
      "data_contratto": "DD/MM/YYYY",
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
"""


# ════════════════════════════════════════════════════════════════════════════
#  Ollama — local LLM call (no internet required)
# ════════════════════════════════════════════════════════════════════════════

def _detect_ollama_host() -> str:
    """
    Auto-detect which host Ollama is actually running on.
    Windows sometimes uses 127.0.0.1 instead of localhost.
    Returns the first working host, or falls back to OLLAMA_HOST from env.
    """
    candidates = [
        OLLAMA_HOST,                    # from .env (first priority)
        "http://localhost:11434",
        "http://127.0.0.1:11434",
    ]
    for host in candidates:
        try:
            req = urllib.request.Request(f"{host}/api/tags")
            with urllib.request.urlopen(req, timeout=3) as resp:
                if resp.status == 200:
                    if host != OLLAMA_HOST:
                        log.info(f"Ollama auto-detected at {host} (env said {OLLAMA_HOST})")
                    return host
        except Exception:
            continue
    return OLLAMA_HOST  # fallback even if not reachable


# Detect once at startup
_ACTIVE_OLLAMA_HOST = None

def get_ollama_host() -> str:
    global _ACTIVE_OLLAMA_HOST
    if _ACTIVE_OLLAMA_HOST is None:
        _ACTIVE_OLLAMA_HOST = _detect_ollama_host()
    return _ACTIVE_OLLAMA_HOST


def check_ollama_health() -> bool:
    """Check if Ollama server is running — tries multiple hosts."""
    host = get_ollama_host()
    try:
        req = urllib.request.Request(f"{host}/api/tags")
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status == 200
    except Exception:
        # Reset so next call re-detects
        global _ACTIVE_OLLAMA_HOST
        _ACTIVE_OLLAMA_HOST = None
        return False


def call_ollama(prompt: str, system: str) -> str:
    """
    Call local Ollama server.
    No internet needed — runs 100% on your machine.
    """
    host = get_ollama_host()
    payload = json.dumps({
        "model":  OLLAMA_MODEL,
        "system": system,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature":    0.0,   # deterministic — best for extraction
            "num_predict":    4096,  # max output tokens
            "num_ctx":        32768, # context window (handles long PDFs)
            "top_p":          1.0,
            "repeat_penalty": 1.1,
        },
    }).encode("utf-8")

    req = urllib.request.Request(
        f"{host}/api/generate",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=OLLAMA_TIMEOUT) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            return body.get("response", "")
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"Cannot reach Ollama at {host}. "
            f"Is Ollama running? Run: ollama serve\n"
            f"Error: {e}"
        )


def safe_json(text: str) -> dict:
    """Try multiple strategies to parse JSON from LLM response."""
    # Strategy 1: direct parse
    try:
        return json.loads(text.strip())
    except Exception:
        pass

    # Strategy 2: extract from markdown code block
    try:
        match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
        if match:
            return json.loads(match.group(1))
    except Exception:
        pass

    # Strategy 3: find first { ... } block
    try:
        match = re.search(r"\{.*\}", text, re.DOTALL)
        if match:
            return json.loads(match.group(0))
    except Exception:
        pass

    raise ValueError(f"Cannot parse JSON from LLM response: {text[:500]}")


# ════════════════════════════════════════════════════════════════════════════
#  PDF Extraction
# ════════════════════════════════════════════════════════════════════════════

def extract_pdf_text(pdf_path: str) -> str:
    """
    Smart extraction: only keep pages with actual contract data.
    Skips boilerplate legal pages (articles, glossary, privacy).
    """
    KEY_TERMS = [
        "corrispettivo", "durata", "collaboratore", "documento n",
        "luogo e data", "conto metallo deluxe", "importo previsto",
        "frequenza", "semestrale", "trimestrale", "mensile",
        "contraente", "cognome e nome", "data operazione",
        "commissioni", "valorizzazione", "informazioni di contratto",
    ]
    relevant_pages = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        log.info(f"PDF has {total} pages — scanning for relevant pages...")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            text_lower = text.lower()
            is_early = i < 5
            has_key_term = any(term in text_lower for term in KEY_TERMS)
            if is_early or has_key_term:
                relevant_pages.append(f"=== PAGE {i+1} ===\n{text}")
                log.info(f"  Page {i+1} included")
            else:
                log.info(f"  Page {i+1} skipped (boilerplate)")
    result = "\n\n".join(relevant_pages)
    log.info(f"Smart extract: {len(relevant_pages)}/{total} pages, {len(result)} chars")
    return result


# def ai_extract(pdf_text: str) -> dict:
#     """Send PDF text to local Ollama and get structured JSON back."""
#     if len(pdf_text) > 15000:
#         log.info(f"Text capped: {len(pdf_text)} -> 15000 chars")
#         pdf_text = pdf_text[:15000]

#     prompt = (
#         "Extract contract data from this Italian gold investment PDF document.\n"
#         "Read the EXACT values from these locations:\n"
#         "- Corrispettivo field = importo_contratto\n"
#         "- Article 2 percentage = commissioni_annuale (Plus=6.00, Top=10.00)\n"
#         "- Article 4.5 frequency = frequenza_pagamento\n"
#         "- Luogo e data = data_contratto\n"
#         "- Durata field = durata_mesi\n\n"
#         f"{pdf_text}\n\n"
#         "Return ONLY the JSON object, no other text."
#     )

#     raw = call_ollama(prompt, SYSTEM_PROMPT)
#     log.info(f"Ollama raw response (first 500 chars): {raw[:500]}")
#     return safe_json(raw)

def ai_extract(pdf_text: str) -> dict:
    """Send PDF text to local Ollama and get structured JSON back."""
    # REMOVED the 15000 cap — qwen2.5:7b supports 32k context
    # Previous code was: if len(pdf_text) > 15000: pdf_text = pdf_text[:15000]

    prompt = (
        "Extract contract data from this Italian gold investment PDF document.\n"
        "Read the EXACT values from these locations:\n"
        "- Corrispettivo field = importo_contratto\n"
        "- Article 2 percentage = commissioni_annuale (Plus=6.00, Top=10.00)\n"
        "- Article 4.5 frequency = frequenza_pagamento\n"
        "- Luogo e data = data_contratto\n"
        "- Durata field = durata_mesi\n\n"
        f"{pdf_text}\n\n"
        "Return ONLY the JSON object."
    )

    raw = call_ollama(prompt, SYSTEM_PROMPT)
    log.info(f"Ollama raw response (first 500 chars): {raw[:500]}")
    return safe_json(raw)


# ════════════════════════════════════════════════════════════════════════════
#  Amount Parser
# ════════════════════════════════════════════════════════════════════════════

def _parse_amount(s: str) -> float:
    """Handles Italian (10.000,00) and English (10000.00) number formats."""
    s = str(s).replace("€", "").replace("$", "").strip()
    if not s:
        return 0.0
    if "." in s and "," in s:
        if s.index(",") > s.index("."):   # Italian: 10.000,00
            s = s.replace(".", "").replace(",", ".")
        else:                              # English: 10,000.00
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
#  Payment Schedule
# ════════════════════════════════════════════════════════════════════════════

MONTHS_IT = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"
]


def fix_schedule(contract: dict) -> list:
    date_str = contract.get("data_contratto", "")
    dt = datetime.strptime(date_str, "%d/%m/%Y")

    durata_mesi  = int(float(str(contract.get("durata_mesi", "")).replace(",", ".")))
    importo      = _parse_amount(str(contract.get("importo_contratto", "")))
    comm_annuale = _parse_amount(str(contract.get("commissioni_annuale", "")))
    pct_periodo  = _parse_amount(str(contract.get("percentuale_per_periodo", "")))
    frequenza    = str(contract.get("frequenza_pagamento", "")).lower()

    if "mens" in frequenza:
        payment_interval = 1
    elif "trim" in frequenza:
        payment_interval = 3
    elif "ann" in frequenza:
        payment_interval = 12
    else:
        payment_interval = 6

    if pct_periodo == 0 and comm_annuale > 0:
        pct_periodo = comm_annuale / (12 / payment_interval)
    amount_per_period = importo * (pct_periodo / 100)

    payment_month_keys = set()
    payment_date_map   = {}
    k = 1
    while True:
        pay_dt       = dt + relativedelta(months=k * payment_interval)
        contract_end = dt + relativedelta(months=durata_mesi)
        if pay_dt > contract_end:
            break
        last_day = calendar.monthrange(pay_dt.year, pay_dt.month)[1]
        key = (pay_dt.year, pay_dt.month)
        payment_month_keys.add(key)
        payment_date_map[key] = f"{pay_dt.month}/{last_day:02d}/{pay_dt.year}"
        k += 1

    if payment_month_keys:
        last_pay_dt = max(datetime(y, m, 1) for y, m in payment_month_keys)
        total_rows = (
            (last_pay_dt.year - dt.year) * 12
            + (last_pay_dt.month - dt.month) + 1
        )
    else:
        total_rows = durata_mesi

    schedule = []
    for i in range(total_rows):
        current = dt + relativedelta(months=i)
        mese    = MONTHS_IT[current.month - 1]
        key     = (current.year, current.month)
        if key in payment_month_keys:
            pay_date = payment_date_map[key]
            amt      = amount_per_period
            schedule.append({
                "mese":                       mese,
                "data_pagamento_commissioni": pay_date,
                "management_mensile":         f"{amt:.2f}",
                "_amount":                    amt,
            })
    return schedule


def enrich_contracts(data: dict) -> dict:
    for c in data.get("contracts", []):
        c["_importo"]      = _parse_amount(str(c.get("importo_contratto", "0")))
        c["_comm_annuale"] = _parse_amount(str(c.get("commissioni_annuale", "0")))
        c["_pct_periodo"]  = _parse_amount(str(c.get("percentuale_per_periodo", "0")))
        c["commissioni"]   = str(c["_comm_annuale"])
        c["_comm"]         = c["_comm_annuale"]
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
    r.font.bold  = True
    r.font.size  = Pt(size)
    r.font.name  = "Calibri"
    r.font.color.rgb = _WHITE


def _data_cell(cell, text: str, bold=False, center=True, bg="FFFFFF", color=None):
    _shd(cell, bg)
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(str(text) if text else "")
    r.font.bold  = bold
    r.font.size  = Pt(8)
    r.font.name  = "Calibri"
    r.font.color.rgb = color if color else _DARK


def _fmt_eur(amount: float) -> str:
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
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.name = "Calibri"
    r.font.color.rgb = _NAVY

    s  = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = s.add_run(
        f"Doc: {data.get('document_number','N/A')}  |  "
        f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    sr.font.size = Pt(8)
    sr.font.name = "Calibri"
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
        lr.font.bold = True
        lr.font.size = Pt(9)
        lr.font.name = "Calibri"
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
            _data_cell(row.cells[6], pay, bg=bg)

            amt = sched.get("_amount", 0.0)
            if not amt:
                amt = _parse_amount(sched.get("management_mensile", ""))
            total_paid += amt
            amt_fmt = _fmt_eur(amt) if amt else ""
            _data_cell(row.cells[7], amt_fmt, bold=bool(amt), bg=bg,
                       color=_NAVY if amt else None)

        total_row = tbl.rows[1 + n_rows]
        _shd(total_row.cells[0], "E8F0FE")
        total_row.cells[0].merge(total_row.cells[5])
        p = total_row.cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r2 = p.add_run(f"TOTALE COMMISSIONI — {nome}")
        r2.font.bold = True
        r2.font.size = Pt(9)
        r2.font.name = "Calibri"
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
        nr.font.size   = Pt(7)
        nr.font.italic = True
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

_FONT_HDR   = Font(name="Calibri", bold=True,  color="FFFFFF",  size=9)
_FONT_TIPO  = Font(name="Calibri", bold=True,  color="1F497D",  size=8)
_FONT_DATA  = Font(name="Calibri", size=8,     color="262626")
_FONT_AMT   = Font(name="Calibri", bold=True,  size=8,  color="1F497D")
_FONT_TOTAL = Font(name="Calibri", bold=True,  size=9,  color="1F497D")
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

            _cell(ws, r, 1, tipo if ri == 0 else "",
                  font=_FONT_TIPO if ri == 0 else _FONT_DATA,
                  fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)
            _cell(ws, r, 2, nome if ri == 0 else "",
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)
            _cell(ws, r, 3, datec if ri == 0 else "",
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)
            _cell(ws, r, 4, importo_raw if ri == 0 else None,
                  font=Font(name="Calibri", bold=(ri == 0), size=8,
                            color="1F497D" if ri == 0 else "262626"),
                  fill=stripe, align=_ALIGN_R, border=_BORDER_THIN,
                  num_fmt='#,##0.00 "€"')
            comm_label = f"{comm_annuale:.2f}% p.a. ({piano})" if ri == 0 else None
            _cell(ws, r, 5, comm_label,
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)
            _cell(ws, r, 6, sched.get("mese", ""),
                  font=Font(name="Calibri", bold=True, size=8),
                  fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)
            pay = sched.get("data_pagamento_commissioni", "")
            _cell(ws, r, 7, pay if pay else None,
                  font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)
            amt = sched.get("_amount", 0.0)
            if not amt:
                amt = _parse_amount(sched.get("management_mensile", ""))
            _cell(ws, r, 8, amt if amt else None,
                  font=_FONT_AMT if amt else _FONT_DATA,
                  fill=stripe, align=_ALIGN_R, border=_BORDER_THIN,
                  num_fmt='#,##0.00 "€"')

        current_row += len(schedule)

        total_amt = sum(s.get("_amount", 0.0) for s in schedule)
        ws.row_dimensions[current_row].height = 20
        ws.merge_cells(f"A{current_row}:F{current_row}")
        _cell(ws, current_row, 1, f"TOTALE COMMISSIONI — {nome}",
              font=_FONT_TOTAL, fill=_FILL_TOTAL,
              align=_ALIGN_L, border=_BORDER_THICK)
        for col in range(2, 7):
            _cell(ws, current_row, col, None, fill=_FILL_TOTAL, border=_BORDER_THICK)
        _cell(ws, current_row, 7, "TOTALE",
              font=_FONT_TOTAL, fill=_FILL_TOTAL,
              align=_ALIGN_C, border=_BORDER_THICK)
        _cell(ws, current_row, 8, total_amt,
              font=Font(name="Calibri", bold=True, size=10, color="059669"),
              fill=_FILL_TOTAL, align=_ALIGN_R, border=_BORDER_THICK,
              num_fmt='#,##0.00 "€"')
        current_row += 2

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    for col, w in zip("ABCDE", [30, 20, 18, 18, 20]):
        ws2.column_dimensions[col].width = w

    ws2.merge_cells("A1:E1")
    _cell(ws2, 1, 1, "CONTRACT SUMMARY",
          font=_FONT_TITLE, fill=_FILL_NAVY, align=_ALIGN_C)
    ws2.row_dimensions[1].height = 28

    for ci2, hdr in enumerate(["Cliente/Segnalatore", "Importo", "Comm % Annuo",
                                "Annual Return", "Total Paid"], 1):
        _cell(ws2, 2, ci2, hdr, font=_FONT_HDR, fill=_FILL_NAVY,
              align=_ALIGN_C, border=_BORDER_THIN)
    ws2.row_dimensions[2].height = 22

    for ri, c in enumerate(contracts, start=3):
        importo_r    = c.get("_importo", 0.0)
        comm_a       = c.get("_comm_annuale", 0.0)
        annual_return = importo_r * (comm_a / 100)
        paid = sum(s.get("_amount", 0.0) for s in c.get("payment_schedule", []))
        stripe2 = _FILL_STRIPE if ri % 2 else _FILL_WHITE
        _cell(ws2, ri, 1, c.get("nome_cliente_segnalatore", ""),
              font=_FONT_DATA, fill=stripe2, align=_ALIGN_L, border=_BORDER_THIN)
        _cell(ws2, ri, 2, importo_r, font=_FONT_AMT, fill=stripe2,
              align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
        _cell(ws2, ri, 3, comm_a / 100, font=_FONT_DATA, fill=stripe2,
              align=_ALIGN_C, border=_BORDER_THIN, num_fmt='0.00%')
        _cell(ws2, ri, 4, annual_return, font=_FONT_AMT, fill=stripe2,
              align=_ALIGN_R, border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
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
    ollama_ok = check_ollama_health()
    return {
        "status":             "ok" if ollama_ok else "degraded",
        "ollama_ready":       ollama_ok,
        "ollama_host":        get_ollama_host(),
        "ollama_host_config": OLLAMA_HOST,
        "model":              OLLAMA_MODEL,
        "mode":               "offline/self-hosted",
        "timestamp":          datetime.now().isoformat(),
    }


@app.get("/api/models")
def list_models():
    """List all models available in local Ollama."""
    try:
        req = urllib.request.Request(f"{OLLAMA_HOST}/api/tags")
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            models = [m["name"] for m in data.get("models", [])]
            return {"models": models, "current": OLLAMA_MODEL}
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"Ollama not reachable: {e}")


@app.post("/api/upload")
async def upload_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files accepted")

    if not check_ollama_health():
        raise HTTPException(
            status_code=503,
            detail=f"Ollama is not running at {OLLAMA_HOST}. Run: ollama serve"
        )

    job_id   = str(uuid.uuid4())[:8]
    pdf_path = UPLOAD_DIR / f"{job_id}_{file.filename}"

    content = await file.read()
    with open(pdf_path, "wb") as f:
        f.write(content)
    log.info(f"Saved PDF: {pdf_path} ({pdf_path.stat().st_size} bytes)")

    try:
        pdf_text  = extract_pdf_text(str(pdf_path))
        extracted = ai_extract(pdf_text)
        log.info(f"Extracted {len(extracted.get('contracts', []))} contracts")

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
        generate_docx(data, str(OUTPUT_DIR / f"{req.job_id}_commissioni.docx"))
        result["docx_ready"] = True
    except Exception as e:
        result["docx_error"] = str(e)
        result["docx_ready"] = False
    try:
        generate_excel(data, str(OUTPUT_DIR / f"{req.job_id}_commissioni.xlsx"))
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
            "total_payment_rows":       len(c.get("payment_schedule", [])),
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
        raise HTTPException(status_code=400, detail="fmt must be 'docx' or 'excel'")
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not generated yet")
    return FileResponse(path=str(path), filename=name, media_type=media)


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    log.info(f"Starting MetalloDoc Offline server on port {PORT}")
    log.info(f"Ollama: {OLLAMA_HOST}  Model: {OLLAMA_MODEL}")
    uvicorn.run("main:app", host="0.0.0.0", port=PORT, reload=False)





















#----clientine kaanicha code using claude ai--------------------------------------
# """
# # main.py — PDF → DOCX + Excel Extractor (FastAPI version)
# # FIXED: dynamic extraction, no hardcoded commission rates
# """
# #---clienntin kaanicha code ----

# import os
# import re
# import json
# import uuid
# import logging
# import traceback
# import urllib.request
# from datetime import datetime
# from pathlib import Path

# # ── Load .env ────────────────────────────────────────────────────────────────
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

# # ── Logging ───────────────────────────────────────────────────────────────────
# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s %(levelname)s %(message)s",
#     datefmt="%Y-%m-%d %H:%M:%S",
# )
# log = logging.getLogger(__name__)

# # ── Config ────────────────────────────────────────────────────────────────────
# # OPENAI_API_KEY    = os.environ.get("OPENAI_API_KEY", "")
# # OPENAI_MODEL      = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
# OPENAI_MAX_TOKENS = int(os.environ.get("OPENAI_MAX_TOKENS", "4096"))
# PORT              = int(os.environ.get("PORT", "8000"))

# ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
# ANTHROPIC_MODEL   = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-20250514")

# log.info(f"Model  : {ANTHROPIC_MODEL}")
# log.info(f"AI Key : {'SET' if ANTHROPIC_API_KEY else 'NOT SET'}")

# BASE_DIR   = Path(__file__).parent
# UPLOAD_DIR = BASE_DIR / "uploads"
# OUTPUT_DIR = BASE_DIR / "outputs"
# UPLOAD_DIR.mkdir(exist_ok=True)
# OUTPUT_DIR.mkdir(exist_ok=True)

# # ── FastAPI ───────────────────────────────────────────────────────────────────
# from fastapi import FastAPI, File, UploadFile, HTTPException
# from fastapi.responses import FileResponse
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
# from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
# from openpyxl.utils import get_column_letter

# app = FastAPI(title="MetalloDoc API", version="2.0.0")
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=False,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# class GenerateRequest(BaseModel):
#     job_id: str
#     data: Optional[Any] = None


# # ════════════════════════════════════════════════════════════════════════════
# #  SYSTEM PROMPT — fully dynamic, no hardcoded rates
# # ════════════════════════════════════════════════════════════════════════════

# SYSTEM_PROMPT = """\
# You are a financial document data extractor for Italian gold investment contracts
# (Conto Metallo Deluxe, Mac & Ro Srl). Extract data and return ONLY valid JSON.

# SEARCH THE ENTIRE DOCUMENT CAREFULLY FOR THESE FIELDS:

# 1. importo_contratto
#    - Search for "Corrispettivo" field
#    - Also check "Importo previsto di investimento" field
#    - Return numeric string only, no € symbol
#    - Example: "10000.00"

# 2. commissioni_annuale
#    - Search Article 2 for percentage
#    - Plus plan = "6.00", Top plan = "10.00"
#    - Return numeric string only

# 3. frequenza_pagamento
#    - Search Article 4 section 5
#    - "semestrale" = every 6 months
#    - "trimestrale" = every 3 months  
#    - "mensile" = every month
#    - Return exact word: "semestrale" or "trimestrale" or "mensile"

# 4. percentuale_per_periodo
#    - semestrale: commissioni_annuale / 2
#    - trimestrale: commissioni_annuale / 4
#    - mensile: commissioni_annuale / 12
#    - Return numeric string only

# 5. data_contratto
#    - Search "Luogo e data" field
#    - Format: DD/MM/YYYY
#    - Example: "2026-01-13" → return "13/01/2026"

# 6. durata_mesi
#    - Search "Durata" field
#    - Return number only: "12" or "24" or "36" or "48" or "60"

# 7. piano
#    - If "Conto Metallo Deluxe Plus" is checked → "Plus"
#    - If "Conto Metallo Deluxe Top" is checked → "Top"

# 8. nome_cliente_segnalatore
#    - Format: "CLIENT_FULLNAME/AGENT_FULLNAME"
#    - Client = CONTRAENTE PERSONA FISICA Cognome e Nome
#    - Agent = COLLABORATORE Cognome e Nome
#    - Example: "QUARANTA PIERPAOLO/Plaitano Dario"

# 9. document_number
#    - Search "Documento n°" field
#    - Example: "CMP-000018.01.2026"

# IMPORTANT RULES:
# - Search the ENTIRE document, not just first pages
# - Extract ONLY values physically written in the PDF
# - NEVER invent or guess any value
# - If a field is not found, use empty string ""
# - Return ONLY raw JSON, no markdown, no explanation

# Return ONLY this JSON:
# {
#   "contracts": [
#     {
#       "tipo_contratto": "Conto Metallo Deluxe",
#       "nome_cliente_segnalatore": "CLIENT/AGENT",
#       "data_contratto": "DD/MM/YYYY",
#       "importo_contratto": "10000.00",
#       "commissioni_annuale": "6.00",
#       "frequenza_pagamento": "semestrale",
#       "percentuale_per_periodo": "3.00",
#       "durata_mesi": "60",
#       "piano": "Plus"
#     }
#   ],
#   "document_number": "CMP-000018.01.2026",
#   "extraction_notes": "Extracted from PDF"
# }
# """


# # ════════════════════════════════════════════════════════════════════════════
# #  PDF + AI extraction
# # ════════════════════════════════════════════════════════════════════════════

# def extract_pdf_text(pdf_path: str) -> str:
#     pages = []
#     with pdfplumber.open(pdf_path) as pdf:
#         for i, page in enumerate(pdf.pages):
#             text = page.extract_text() or ""
#             pages.append(f"=== PAGE {i+1} ===\n{text}")
#     return "\n\n".join(pages)


# def call_claude(prompt: str, system: str) -> str:
#     payload = json.dumps({
#         "model":      ANTHROPIC_MODEL,
#         "max_tokens": 8096,        # maximum tokens for full response
#         "system":     system,
#         "messages": [
#             {"role": "user", "content": prompt},
#         ],
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
#     with urllib.request.urlopen(req, timeout=300) as resp:
#         body = json.loads(resp.read().decode("utf-8"))
#         return body["content"][0]["text"]


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


# def ai_extract(pdf_text: str) -> dict:
#     prompt = (
#         "Extract contract data from this Italian gold investment PDF.\n"
#         "Read the EXACT values from these locations:\n"
#         "- Corrispettivo field = importo_contratto\n"
#         "- Article 2 percentage = commissioni_annuale (Plus=6.00, Top=10.00)\n"
#         "- Article 4.5 frequency = frequenza_pagamento\n"
#         "- Luogo e data = data_contratto\n"
#         "- Durata field = durata_mesi\n\n"
#         f"{pdf_text}\n\n"          # NO limit — full PDF text
#         "Return ONLY the JSON object."
#     )
#     raw = call_claude(prompt, SYSTEM_PROMPT)
#     log.info(f"Claude raw (first 500): {raw[:500]}")
#     return safe_json(raw)


# # ════════════════════════════════════════════════════════════════════════════
# #  Amount parser — handles Italian and English number formats
# # ════════════════════════════════════════════════════════════════════════════

# def _parse_amount(s: str) -> float:
#     """
#     Handles: "10000", "10.000,00", "10,000.00", "10000.00"
#     Returns float, 0.0 on failure — never a hardcoded fallback
#     """
#     s = str(s).replace("€", "").replace("$", "").strip()
#     if not s:
#         return 0.0
#     # Italian format: 10.000,00
#     if "." in s and "," in s:
#         if s.index(",") > s.index("."):
#             s = s.replace(".", "").replace(",", ".")
#         else:
#             s = s.replace(",", "")
#     elif "," in s:
#         parts = s.split(",")
#         if len(parts) == 2 and len(parts[1]) <= 2:
#             s = s.replace(",", ".")
#         else:
#             s = s.replace(",", "")
#     try:
#         return float(re.sub(r"[^\d.]", "", s))
#     except Exception:
#         return 0.0


# # ════════════════════════════════════════════════════════════════════════════
# #  Payment schedule — 100% dynamic, zero hardcoded values
# # ════════════════════════════════════════════════════════════════════════════

# MONTHS_IT    = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
#                 "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]
# MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun",
#                 "Jul","Aug","Sep","Oct","Nov","Dec"]


# # import calendar

# # def fix_schedule(contract: dict) -> list:
# #     date_str = contract.get("data_contratto", "")
# #     try:
# #         dt      = datetime.strptime(date_str, "%d/%m/%Y")
# #         start_m = dt.month - 1
# #         start_y = dt.year
# #     except Exception:
# #         start_m, start_y = 0, 2026

# #     durata_mesi = int(float(str(contract.get("durata_mesi", "12")).replace(",",".")))
# #     if durata_mesi <= 0:
# #         durata_mesi = 12

# #     importo      = _parse_amount(str(contract.get("importo_contratto", "0")))
# #     comm_annuale = _parse_amount(str(contract.get("commissioni_annuale", "0")))
# #     pct_periodo  = _parse_amount(str(contract.get("percentuale_per_periodo", "0")))
# #     frequenza    = str(contract.get("frequenza_pagamento", "semestrale")).lower()

# #     if "mens" in frequenza:
# #         payment_interval = 1
# #     elif "trim" in frequenza:
# #         payment_interval = 3
# #     elif "ann" in frequenza:
# #         payment_interval = 12
# #     else:
# #         payment_interval = 6

# #     if pct_periodo == 0 and comm_annuale > 0:
# #         pct_periodo = comm_annuale / (12 / payment_interval)

# #     amount_per_period = importo * (pct_periodo / 100)

# #     # Grace period: 40 days from contract date
# #     contract_day = dt.day if date_str else 13
# #     days_remaining = calendar.monthrange(start_y, start_m + 1)[1] - contract_day
# #     grace_months = 1 if days_remaining >= 40 else 2

# #     schedule = []
# #     active_month_counter = 0

# #     for i in range(durata_mesi):
# #         total_months = start_m + i
# #         month_idx    = total_months % 12
# #         year         = start_y + total_months // 12
# #         mese         = f"{MONTHS_IT[month_idx]} {year}"

# #         if i < grace_months:
# #             schedule.append({
# #                 "mese":                       mese,
# #                 "data_pagamento_commissioni": "",
# #                 "management_mensile":         "",
# #                 "_amount":                    0.0,
# #             })
# #         else:
# #             active_month_counter += 1
# #             is_payment = (active_month_counter % payment_interval == 0)

# #             if is_payment:
# #                 last_day = calendar.monthrange(year, month_idx + 1)[1]
# #                 pay_date = f"{last_day:02d}/{month_idx+1:02d}/{year}"
# #                 amt      = amount_per_period
# #             else:
# #                 pay_date = ""
# #                 amt      = 0.0

# #             schedule.append({
# #                 "mese":                       mese,
# #                 "data_pagamento_commissioni": pay_date,
# #                 "management_mensile":         f"{amt:.2f}" if amt else "",
# #                 "_amount":                    amt,
# #             })

# #     return schedule





# import calendar
# from datetime import datetime
# from dateutil.relativedelta import relativedelta   # pip install python-dateutil
 
# MONTHS_IT = [
#     "Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
#     "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"
# ]
 
 
# def _parse_amount(s: str) -> float:
#     """Handles Italian (10.000,00) and English (10000.00) number formats."""
#     import re
#     s = str(s).replace("€","").replace("$","").strip()
#     if not s:
#         return 0.0
#     if "." in s and "," in s:
#         if s.index(",") > s.index("."):          # Italian: 10.000,00
#             s = s.replace(".", "").replace(",", ".")
#         else:                                      # English: 10,000.00
#             s = s.replace(",", "")
#     elif "," in s:
#         parts = s.split(",")
#         if len(parts) == 2 and len(parts[1]) <= 2:
#             s = s.replace(",", ".")
#         else:
#             s = s.replace(",", "")
#     try:
#         return float(re.sub(r"[^\d.]", "", s))
#     except Exception:
#         return 0.0
 
 
# def fix_schedule(contract: dict) -> list:
#     # ── 1. Parse contract date — no fallback, crash if missing ───────────────
#     date_str = contract.get("data_contratto", "")
#     dt = datetime.strptime(date_str, "%d/%m/%Y")

#     # ── 2. Contract parameters — all from PDF ────────────────────────────────
#     durata_mesi = int(float(str(contract.get("durata_mesi", "")).replace(",", ".")))
#     importo      = _parse_amount(str(contract.get("importo_contratto", "")))
#     comm_annuale = _parse_amount(str(contract.get("commissioni_annuale", "")))
#     pct_periodo  = _parse_amount(str(contract.get("percentuale_per_periodo", "")))
#     frequenza    = str(contract.get("frequenza_pagamento", "")).lower()

#     # ── 3. Payment interval from PDF field ───────────────────────────────────
#     if "mens" in frequenza:
#         payment_interval = 1
#     elif "trim" in frequenza:
#         payment_interval = 3
#     elif "ann" in frequenza:
#         payment_interval = 12
#     else:  # semestrale
#         payment_interval = 6

#     # ── 4. Amount per period ─────────────────────────────────────────────────
#     if pct_periodo == 0 and comm_annuale > 0:
#         pct_periodo = comm_annuale / (12 / payment_interval)
#     amount_per_period = importo * (pct_periodo / 100)

#     # ── 5. Build payment dates: start_date + k*interval months ───────────────
#     #    Per PDF: "Le date di versamento delle valorizzazioni sono calcolate
#     #    a partire dalla data di decorrenza del Contratto"
#     #    → payment 1 = dt + 1*interval months, payment 2 = dt + 2*interval, ...
#     payment_month_keys = set()
#     payment_date_map   = {}
#     k = 1
#     while True:
#         pay_dt = dt + relativedelta(months=k * payment_interval)
#         # Stop when beyond contract end
#         contract_end = dt + relativedelta(months=durata_mesi)
#         if pay_dt > contract_end:
#             break
#         last_day = calendar.monthrange(pay_dt.year, pay_dt.month)[1]
#         key = (pay_dt.year, pay_dt.month)
#         payment_month_keys.add(key)
#         payment_date_map[key] = f"{pay_dt.month}/{last_day:02d}/{pay_dt.year}"
#         k += 1

#     # ── 6. Rows: from contract month through LAST payment month ──────────────
#     #    (extends beyond durata_mesi if last payment lands there)
#     if payment_month_keys:
#         last_pay_dt = max(datetime(y, m, 1) for y, m in payment_month_keys)
#         total_rows = (
#             (last_pay_dt.year - dt.year) * 12
#             + (last_pay_dt.month - dt.month)
#             + 1
#         )
#     else:
#         total_rows = durata_mesi

#     # ── 7. Build schedule rows ────────────────────────────────────────────────
#     schedule = []
#     for i in range(total_rows):
#         current  = dt + relativedelta(months=i)
#         mese     = f"{MONTHS_IT[current.month - 1]}"  # ← REMOVED year
#         key      = (current.year, current.month)

#         if key in payment_month_keys:          # ← ONLY add row if payment exists
#             pay_date = payment_date_map[key]
#             amt      = amount_per_period
#             schedule.append({
#                 "mese":                       mese,
#                 "data_pagamento_commissioni": pay_date,
#                 "management_mensile":         f"{amt:.2f}" if amt else "",
#                 "_amount":                    amt,
#             })
#         # ← non-payment months are simply skipped

#     return schedule
# def enrich_contracts(data: dict) -> dict:
#     """Add calculated fields to each contract — no hardcoded defaults."""
#     for c in data.get("contracts", []):
#         c["_importo"]       = _parse_amount(str(c.get("importo_contratto", "0")))
#         c["_comm_annuale"]  = _parse_amount(str(c.get("commissioni_annuale", "0")))
#         c["_pct_periodo"]   = _parse_amount(str(c.get("percentuale_per_periodo", "0")))

#         # Keep "commissioni" field for display in table header
#         c["commissioni"]    = str(c["_comm_annuale"])
#         c["_comm"]          = c["_comm_annuale"]

#         c["payment_schedule"] = fix_schedule(c)

#     return data
 


# # ════════════════════════════════════════════════════════════════════════════
# #  DOCX Generator
# # ════════════════════════════════════════════════════════════════════════════

# _NAVY  = RGBColor(0x1F, 0x49, 0x7D)
# _GOLD  = RGBColor(0xC0, 0x9B, 0x3A)
# _WHITE = RGBColor(0xFF, 0xFF, 0xFF)
# _DARK  = RGBColor(0x26, 0x26, 0x26)


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


# def _fmt_eur(amount: float) -> str:
#     """Format as Italian currency: 10.000,00 €"""
#     return f"{amount:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


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

#     s  = doc.add_paragraph()
#     s.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     sr = s.add_run(
#         f"Doc: {data.get('document_number','N/A')}  |  "
#         f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
#     )
#     sr.font.size, sr.font.name = Pt(8), "Calibri"
#     sr.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
#     doc.add_paragraph()

#     contracts = data.get("contracts", [])
#     if not contracts:
#         doc.add_paragraph("No contracts found in PDF.")
#         doc.save(out_path)
#         return

#     for ci, contract in enumerate(contracts):
#         lbl = doc.add_paragraph()
#         lr  = lbl.add_run(f"Contract #{ci+1}  —  {contract.get('nome_cliente_segnalatore','')}")
#         lr.font.bold, lr.font.size, lr.font.name = True, Pt(9), "Calibri"
#         lr.font.color.rgb = _GOLD

#         schedule = contract.get("payment_schedule", [])
#         n_rows   = max(len(schedule), 1)

#         tbl = doc.add_table(rows=1 + n_rows + 1, cols=8)
#         tbl.style     = "Table Grid"
#         tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

#         col_widths = [3.0, 4.8, 2.6, 3.0, 2.4, 2.6, 3.8, 3.4]
#         for row in tbl.rows:
#             for i, cell in enumerate(row.cells):
#                 cell.width = Cm(col_widths[i])

#         headers = [
#             "Tipo Contratto", "Nome Cliente/Segnalatore", "Data Contratto",
#             "Importo Contratto", "Commissioni\n(% annuo)", "Mese",
#             "Data Pagamento\nCommissioni", "Management\nMensile"
#         ]
#         for i, txt in enumerate(headers):
#             _hdr_cell(tbl.rows[0].cells[i], txt)

#         importo_raw  = contract.get("_importo", 0.0)
#         comm_annuale = contract.get("_comm_annuale", 0.0)
#         tipo         = contract.get("tipo_contratto", "Conto Metallo Deluxe")
#         nome         = contract.get("nome_cliente_segnalatore", "")
#         data_c       = contract.get("data_contratto", "")
#         piano        = contract.get("piano", "")
#         importo_fmt  = _fmt_eur(importo_raw)
#         # Show annual commission % with piano label
#         comm_display = f"{comm_annuale:.2f}% p.a. ({piano})"

#         total_paid = 0.0
#         for ri, sched in enumerate(schedule):
#             row = tbl.rows[1 + ri]
#             bg  = "FFFFFF" if ri % 2 == 0 else "F4F0E8"

#             _data_cell(row.cells[0], tipo if ri == 0 else "", bold=(ri == 0), center=False, bg=bg)
#             _data_cell(row.cells[1], nome if ri == 0 else "", center=False, bg=bg)
#             _data_cell(row.cells[2], data_c if ri == 0 else "", bg=bg)
#             _data_cell(row.cells[3], importo_fmt if ri == 0 else "", bold=(ri == 0), bg=bg)
#             _data_cell(row.cells[4], comm_display if ri == 0 else "", bg=bg)
#             _data_cell(row.cells[5], sched.get("mese", ""), bold=True, bg=bg)

#             pay = sched.get("data_pagamento_commissioni", "")
#             _data_cell(row.cells[6], pay, bg=bg)

#             amt = sched.get("_amount", 0.0)
#             if not amt:
#                 raw_s = sched.get("management_mensile", "")
#                 amt   = _parse_amount(raw_s) if raw_s else 0.0
#             total_paid += amt
#             amt_fmt = _fmt_eur(amt) if amt else ""
#             _data_cell(row.cells[7], amt_fmt, bold=bool(amt), bg=bg,
#                        color=_NAVY if amt else None)

#         # Total row
#         total_row = tbl.rows[1 + n_rows]
#         _shd(total_row.cells[0], "E8F0FE")
#         total_row.cells[0].merge(total_row.cells[5])
#         p = total_row.cells[0].paragraphs[0]
#         p.alignment = WD_ALIGN_PARAGRAPH.LEFT
#         r2 = p.add_run(f"TOTALE COMMISSIONI — {nome}")
#         r2.font.bold, r2.font.size, r2.font.name = True, Pt(9), "Calibri"
#         r2.font.color.rgb = _NAVY

#         _data_cell(total_row.cells[6], "TOTALE", bold=True, bg="E8F0FE", color=_NAVY)
#         _data_cell(total_row.cells[7], _fmt_eur(total_paid), bold=True, bg="E8F0FE",
#                    color=RGBColor(0x05, 0x96, 0x69))

#         if ci < len(contracts) - 1:
#             doc.add_paragraph()

#     notes = data.get("extraction_notes", "")
#     if notes:
#         doc.add_paragraph()
#         np_ = doc.add_paragraph()
#         nr  = np_.add_run(f"  {notes}")
#         nr.font.size, nr.font.italic = Pt(7), True
#         nr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

#     doc.save(out_path)
#     log.info(f"DOCX saved: {out_path}")


# # ════════════════════════════════════════════════════════════════════════════
# #  Excel Generator
# # ════════════════════════════════════════════════════════════════════════════

# _FILL_NAVY   = PatternFill("solid", fgColor="1F497D")
# _FILL_GOLD   = PatternFill("solid", fgColor="C09B3A")
# _FILL_STRIPE = PatternFill("solid", fgColor="F4F0E8")
# _FILL_WHITE  = PatternFill("solid", fgColor="FFFFFF")
# _FILL_TOTAL  = PatternFill("solid", fgColor="E8F0FE")

# _FONT_HDR   = Font(name="Calibri", bold=True,  color="FFFFFF", size=9)
# _FONT_TIPO  = Font(name="Calibri", bold=True,  color="1F497D", size=8)
# _FONT_DATA  = Font(name="Calibri", size=8,     color="262626")
# _FONT_AMT   = Font(name="Calibri", bold=True,  size=8, color="1F497D")
# _FONT_TOTAL = Font(name="Calibri", bold=True,  size=9, color="1F497D")
# _FONT_TITLE = Font(name="Calibri", bold=True,  size=14, color="1F497D")
# _FONT_SUB   = Font(name="Calibri", size=9,     color="707070")

# _ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
# _ALIGN_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
# _ALIGN_R = Alignment(horizontal="right",  vertical="center")

# _THIN  = Side(style="thin",   color="BFBFBF")
# _THICK = Side(style="medium", color="1F497D")
# _BORDER_THIN  = Border(left=_THIN,  right=_THIN,  top=_THIN,  bottom=_THIN)
# _BORDER_THICK = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)

# HEADERS = [
#     "Tipo Contratto", "Nome Cliente/Segnalatore", "Data Contratto",
#     "Importo Contratto", "Commissioni (% annuo)", "Mese",
#     "Data Pagamento Commissioni", "Management Mensile (€)",
# ]
# COL_W = [22, 30, 14, 18, 16, 14, 24, 22]


# def _cell(ws, row, col, value, font=None, fill=None, align=None, border=None, num_fmt=None):
#     c = ws.cell(row=row, column=col, value=value)
#     if font:    c.font          = font
#     if fill:    c.fill          = fill
#     if align:   c.alignment     = align
#     if border:  c.border        = border
#     if num_fmt: c.number_format = num_fmt
#     return c


# def generate_excel(data: dict, out_path: str):
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Commissioni"
#     ws.freeze_panes = "A5"

#     ws.merge_cells("A1:H1")
#     _cell(ws, 1, 1, "CONTO METALLO DELUXE — GESTIONE COMMISSIONI",
#           font=_FONT_TITLE,
#           fill=PatternFill("solid", fgColor="1F497D"),
#           align=_ALIGN_C)

#     ws.merge_cells("A2:H2")
#     _cell(ws, 2, 1,
#           f"Doc: {data.get('document_number','N/A')}   |   "
#           f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
#           font=_FONT_SUB,
#           fill=PatternFill("solid", fgColor="C09B3A"),
#           align=_ALIGN_C)

#     ws.row_dimensions[3].height = 6

#     for ci, (hdr, w) in enumerate(zip(HEADERS, COL_W), start=1):
#         _cell(ws, 4, ci, hdr,
#               font=_FONT_HDR, fill=_FILL_NAVY,
#               align=_ALIGN_C, border=_BORDER_THIN)
#         ws.column_dimensions[get_column_letter(ci)].width = w

#     ws.row_dimensions[1].height = 28
#     ws.row_dimensions[2].height = 18
#     ws.row_dimensions[4].height = 32

#     current_row = 5
#     contracts   = data.get("contracts", [])

#     for contract in contracts:
#         schedule     = contract.get("payment_schedule", [])
#         importo_raw  = contract.get("_importo", 0.0)
#         comm_annuale = contract.get("_comm_annuale", 0.0)
#         piano        = contract.get("piano", "")
#         tipo         = contract.get("tipo_contratto", "Conto Metallo Deluxe")
#         nome         = contract.get("nome_cliente_segnalatore", "")
#         datec        = contract.get("data_contratto", "")

#         for ri, sched in enumerate(schedule):
#             stripe = _FILL_STRIPE if ri % 2 else _FILL_WHITE
#             r      = current_row + ri
#             ws.row_dimensions[r].height = 18

#             _cell(ws, r, 1,
#                   tipo if ri == 0 else "",
#                   font=_FONT_TIPO if ri == 0 else _FONT_DATA,
#                   fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)

#             _cell(ws, r, 2,
#                   nome if ri == 0 else "",
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_L, border=_BORDER_THIN)

#             _cell(ws, r, 3,
#                   datec if ri == 0 else "",
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             _cell(ws, r, 4,
#                   importo_raw if ri == 0 else None,
#                   font=Font(name="Calibri", bold=(ri==0), size=8,
#                             color="1F497D" if ri==0 else "262626"),
#                   fill=stripe, align=_ALIGN_R, border=_BORDER_THIN,
#                   num_fmt='#,##0.00 "€"')

#             # Commission: show annual % with piano label
#             comm_label = f"{comm_annuale:.2f}% p.a. ({piano})" if ri == 0 else None
#             _cell(ws, r, 5,
#                   comm_label,
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             _cell(ws, r, 6,
#                   sched.get("mese", ""),
#                   font=Font(name="Calibri", bold=True, size=8),
#                   fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             pay = sched.get("data_pagamento_commissioni", "")
#             _cell(ws, r, 7,
#                   pay if pay else None,
#                   font=_FONT_DATA, fill=stripe, align=_ALIGN_C, border=_BORDER_THIN)

#             amt = sched.get("_amount", 0.0)
#             if not amt:
#                 raw_s = sched.get("management_mensile", "")
#                 amt   = _parse_amount(raw_s) if raw_s else 0.0

#             _cell(ws, r, 8,
#                   amt if amt else None,
#                   font=_FONT_AMT if amt else _FONT_DATA,
#                   fill=stripe, align=_ALIGN_R, border=_BORDER_THIN,
#                   num_fmt='#,##0.00 "€"')

#         current_row += len(schedule)

#         # Total row
#         total_amt = sum(s.get("_amount", 0.0) for s in schedule)
#         ws.row_dimensions[current_row].height = 20
#         ws.merge_cells(f"A{current_row}:F{current_row}")
#         _cell(ws, current_row, 1,
#               f"TOTALE COMMISSIONI — {nome}",
#               font=_FONT_TOTAL, fill=_FILL_TOTAL,
#               align=_ALIGN_L, border=_BORDER_THICK)
#         for col in range(2, 7):
#             _cell(ws, current_row, col, None,
#                   fill=_FILL_TOTAL, border=_BORDER_THICK)
#         _cell(ws, current_row, 7,
#               "TOTALE",
#               font=_FONT_TOTAL, fill=_FILL_TOTAL,
#               align=_ALIGN_C, border=_BORDER_THICK)
#         _cell(ws, current_row, 8,
#               total_amt,
#               font=Font(name="Calibri", bold=True, size=10, color="059669"),
#               fill=_FILL_TOTAL, align=_ALIGN_R, border=_BORDER_THICK,
#               num_fmt='#,##0.00 "€"')

#         current_row += 2

#     # ── Summary sheet ─────────────────────────────────────────────────────────
#     ws2 = wb.create_sheet("Summary")
#     for col, w in zip("ABCDE", [30, 20, 18, 18, 20]):
#         ws2.column_dimensions[col].width = w

#     ws2.merge_cells("A1:E1")
#     _cell(ws2, 1, 1, "CONTRACT SUMMARY",
#           font=_FONT_TITLE, fill=_FILL_NAVY, align=_ALIGN_C)
#     ws2.row_dimensions[1].height = 28

#     summary_headers = [
#         "Cliente/Segnalatore", "Importo", "Comm % Annuo",
#         "Annual Return", "Total Paid (Year 1)"
#     ]
#     for ci2, hdr in enumerate(summary_headers, 1):
#         _cell(ws2, 2, ci2, hdr,
#               font=_FONT_HDR, fill=_FILL_NAVY,
#               align=_ALIGN_C, border=_BORDER_THIN)
#     ws2.row_dimensions[2].height = 22

#     for ri, c in enumerate(contracts, start=3):
#         importo_r    = c.get("_importo", 0.0)
#         comm_annuale = c.get("_comm_annuale", 0.0)

#         # Annual return = importo × annual commission %  (correct formula)
#         annual_return = importo_r * (comm_annuale / 100)

#         # Total paid = sum of all payment amounts in the schedule
#         paid = sum(s.get("_amount", 0.0) for s in c.get("payment_schedule", []))

#         stripe2 = _FILL_STRIPE if ri % 2 else _FILL_WHITE

#         _cell(ws2, ri, 1, c.get("nome_cliente_segnalatore", ""),
#               font=_FONT_DATA, fill=stripe2, align=_ALIGN_L, border=_BORDER_THIN)
#         _cell(ws2, ri, 2, importo_r,
#               font=_FONT_AMT, fill=stripe2, align=_ALIGN_R,
#               border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         _cell(ws2, ri, 3, comm_annuale / 100,
#               font=_FONT_DATA, fill=stripe2, align=_ALIGN_C,
#               border=_BORDER_THIN, num_fmt='0.00%')
#         _cell(ws2, ri, 4, annual_return,
#               font=_FONT_AMT, fill=stripe2, align=_ALIGN_R,
#               border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         _cell(ws2, ri, 5, paid,
#               font=Font(name="Calibri", bold=True, size=8, color="059669"),
#               fill=stripe2, align=_ALIGN_R,
#               border=_BORDER_THIN, num_fmt='#,##0.00 "€"')
#         ws2.row_dimensions[ri].height = 18

#     wb.save(out_path)
#     log.info(f"XLSX saved: {out_path}")


# # ════════════════════════════════════════════════════════════════════════════
# #  Helpers
# # ════════════════════════════════════════════════════════════════════════════

# def _load_job_data(job_id: str, body_data: Any) -> dict:
#     if body_data:
#         return enrich_contracts(body_data)
#     path = OUTPUT_DIR / f"{job_id}_data.json"
#     if path.exists():
#         with open(path, encoding="utf-8") as fp:
#             return enrich_contracts(json.load(fp))
#     return None


# # ════════════════════════════════════════════════════════════════════════════
# #  API Routes
# # ════════════════════════════════════════════════════════════════════════════

# @app.get("/api/health")
# def health():
#     return {
#         "status":    "ok",
#         "ai_ready":  bool(ANTHROPIC_API_KEY),
#         "model":     ANTHROPIC_MODEL,
#         "timestamp": datetime.now().isoformat(),
#     }


# @app.post("/api/upload")
# async def upload_pdf(file: UploadFile = File(...)):
#     if not file.filename.lower().endswith(".pdf"):
#         raise HTTPException(status_code=400, detail="Only PDF files accepted")

#     job_id   = str(uuid.uuid4())[:8]
#     pdf_path = UPLOAD_DIR / f"{job_id}_{file.filename}"

#     content = await file.read()
#     with open(pdf_path, "wb") as f:
#         f.write(content)
#     log.info(f"Saved PDF: {pdf_path} ({pdf_path.stat().st_size} bytes)")

#     try:
#         pdf_text = extract_pdf_text(str(pdf_path))
#         log.info(f"Extracted {len(pdf_text)} chars from PDF")

#         if not ANTHROPIC_API_KEY:
#             raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not set")

#         extracted = ai_extract(pdf_text)
#         log.info(f"AI extracted: {json.dumps(extracted, indent=2)[:500]}")

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


# @app.post("/api/generate")
# def gen_document(req: GenerateRequest):
#     return gen_docx(req)


# @app.post("/api/preview")
# def preview(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")
#     return data


# @app.post("/api/debug")
# def debug_extraction(req: GenerateRequest):
#     data = _load_job_data(req.job_id, req.data)
#     if not data:
#         raise HTTPException(status_code=404, detail="No data for job_id")
#     result = []
#     for i, c in enumerate(data.get("contracts", [])):
#         result.append({
#             "contract_number":          i + 1,
#             "tipo_contratto":           c.get("tipo_contratto"),
#             "nome_cliente_segnalatore": c.get("nome_cliente_segnalatore"),
#             "data_contratto":           c.get("data_contratto"),
#             "importo_contratto":        c.get("importo_contratto"),
#             "commissioni_annuale":      c.get("commissioni_annuale"),
#             "frequenza_pagamento":      c.get("frequenza_pagamento"),
#             "percentuale_per_periodo":  c.get("percentuale_per_periodo"),
#             "durata_mesi":              c.get("durata_mesi"),
#             "piano":                    c.get("piano"),
#             "_importo_parsed":          c.get("_importo"),
#             "_comm_annuale_parsed":     c.get("_comm_annuale"),
#             "_pct_periodo_parsed":      c.get("_pct_periodo"),
#             "total_months":             len(c.get("payment_schedule", [])),
#             "payment_schedule":         c.get("payment_schedule", []),
#         })
#     return {"debug": result}


# @app.get("/api/download/{job_id}/{fmt}")
# def download_file(job_id: str, fmt: str):
#     if fmt == "docx":
#         path  = OUTPUT_DIR / f"{job_id}_commissioni.docx"
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
#     log.info(f"Starting server on port {PORT}")
#     uvicorn.run("main:app", host="0.0.0.0", port=PORT, reload=False)