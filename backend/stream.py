#----own self hosted ollama model--------------------

"""
app.py — MetalloDoc Streamlit UI (Offline / Self-Hosted)
"""

import streamlit as st
import requests
import pandas as pd

BASE_URL = "http://localhost:8000"

st.set_page_config(
    page_title="MetalloDoc — Contract Extractor",
    page_icon="📄",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .status-ok   { color: #059669; font-weight: 600; }
    .status-fail { color: #dc2626; font-weight: 600; }
    div[data-testid="stMetricValue"] { font-size: 1.4rem; }
</style>
""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 8])
with col_title:
    st.title("📄 MetalloDoc — Contract Extractor")
    st.caption("Offline · Self-Hosted · 100% Private · Powered by Ollama")

st.divider()


# ── Sidebar: Server Status ────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ System Status")

    if st.button("🔄 Refresh Status", use_container_width=True):
        st.rerun()

    try:
        health = requests.get(f"{BASE_URL}/api/health", timeout=4).json()
        ollama_ok = health.get("ollama_ready", False)

        if ollama_ok:
            st.success("✅ Server: Online")
            st.success("✅ Ollama: Running")
        else:
            st.success("✅ Server: Online")
            st.error("❌ Ollama: Not running")
            st.warning("Run: `ollama serve`")

        st.info(f"🤖 Model: `{health.get('model', 'unknown')}`")
        st.info(f"🔒 Mode: {health.get('mode', 'offline')}")

    except Exception:
        st.error("❌ Server: Offline")
        st.warning("Start server:\n```\npython main.py\n```")

    st.divider()

    # Model selector
    st.subheader("🧠 Model Info")
    st.markdown("""
| Model | RAM | Speed |
|-------|-----|-------|
| `qwen2.5:7b` | 8 GB | ~30s ⭐ |
| `mistral:7b` | 8 GB | ~20s |
| `qwen2.5:14b` | 16 GB | ~60s |
""")

    st.divider()
    st.caption("All data stays on your machine.\nNo internet required.")


# ── STEP 1: Upload PDF ────────────────────────────────────────────────────────
st.header("① Upload & Analyze PDF")

uploaded_file = st.file_uploader(
    "Drop your contract PDF here",
    type=["pdf"],
    help="The PDF will be processed locally — nothing is sent to the internet.",
)

if uploaded_file:
    st.info(f"📎 **{uploaded_file.name}** — {uploaded_file.size / 1024:.1f} KB")

col_btn, col_note = st.columns([2, 5])
with col_btn:
    analyze_clicked = st.button(
        "🚀 Analyze PDF",
        use_container_width=True,
        type="primary",
        disabled=(uploaded_file is None),
    )
with col_note:
    st.caption("⏱️ Local LLM takes ~30–60 seconds. Please wait.")

if analyze_clicked and uploaded_file:
    with st.spinner("🔍 Extracting contract data with local AI... (CPU mode: 10–20 min, please wait)"):

        try:
            files    = {"file": (uploaded_file.name, uploaded_file.getvalue(), "application/pdf")}
            response = requests.post(f"{BASE_URL}/api/upload", files=files, timeout=1800)

        except requests.exceptions.ConnectionError:
            st.error("❌ Cannot connect to server. Make sure `python main.py` is running.")
            st.stop()
        except requests.exceptions.Timeout:
            st.error("❌ Request timed out. Try a smaller PDF or switch to `mistral:7b`.")
            st.stop()

    if response.status_code == 200:
        result = response.json()
        st.session_state["job_id"] = result["job_id"]
        st.session_state["data"]   = result["data"]

        contracts_found = result.get("contracts_found", 0)
        doc_number      = result.get("document_number", "—")

        st.success(f"✅ Extraction complete! Found **{contracts_found}** contract(s)")

        # Summary metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("Contracts Found", contracts_found)
        m2.metric("Document Number", doc_number)
        m3.metric("Job ID", result["job_id"])

        # Raw JSON expander
        with st.expander("🔎 View Raw Extracted JSON", expanded=False):
            st.json(result["data"])

    else:
        st.error(f"❌ Extraction failed: {response.text}")

st.divider()


# ── STEP 2: Preview ───────────────────────────────────────────────────────────
st.header("② Preview Contract Data")

if "job_id" not in st.session_state:
    st.info("Upload and analyze a PDF first.")
else:
    if st.button("🔍 Load Preview Table", use_container_width=False):
        with st.spinner("Loading preview..."):
            resp = requests.post(
                f"{BASE_URL}/api/preview",
                json={"job_id": st.session_state["job_id"]},
                timeout=30,
            )

        if resp.status_code == 200:
            preview = resp.json()
            st.session_state["preview"] = preview
        else:
            st.error(f"Preview failed: {resp.text}")

    if "preview" in st.session_state:
        preview   = st.session_state["preview"]
        contracts = preview.get("contracts", [])

        for i, contract in enumerate(contracts):
            nome  = contract.get("nome_cliente_segnalatore", f"Contract {i+1}")
            piano = contract.get("piano", "")

            with st.expander(f"📋 Contract #{i+1} — {nome} ({piano})", expanded=True):

                # Contract info row
                ci1, ci2, ci3, ci4, ci5 = st.columns(5)
                ci1.metric("Amount", f"€ {contract.get('_importo', 0):,.2f}")
                ci2.metric("Duration", f"{contract.get('durata_mesi', '?')} months")
                ci3.metric("Annual %", f"{contract.get('_comm_annuale', 0):.2f}%")
                ci4.metric("Frequency", contract.get("frequenza_pagamento", "—"))
                ci5.metric("Date", contract.get("data_contratto", "—"))

                # Payment schedule table
                schedule = contract.get("payment_schedule", [])
                if schedule:
                    rows = []
                    total = 0.0
                    for row in schedule:
                        amt = row.get("_amount", 0.0)
                        total += amt
                        rows.append({
                            "Mese":                        row.get("mese", ""),
                            "Data Pagamento Commissioni":  row.get("data_pagamento_commissioni", ""),
                            "Management Mensile (€)":      f"{amt:,.2f}" if amt else "",
                        })

                    df = pd.DataFrame(rows)
                    st.dataframe(
                        df,
                        use_container_width=True,
                        hide_index=True,
                        height=min(400, (len(rows) + 1) * 38),
                    )

                    t1, t2 = st.columns([4, 1])
                    t1.markdown(f"**Total payments: {len(schedule)}**")
                    t2.metric("Total Commissioni", f"€ {total:,.2f}")
                else:
                    st.warning("No payment schedule generated.")

st.divider()


# ── STEP 3: Generate & Download ───────────────────────────────────────────────
st.header("③ Generate & Download")

if "job_id" not in st.session_state:
    st.info("Upload and analyze a PDF first.")
else:
    col_d, col_e, col_b = st.columns(3)

    # ── DOCX ──
    with col_d:
        st.subheader("📄 Word Document")
        st.caption("Professional DOCX with commission table")
        if st.button("Generate DOCX", use_container_width=True, key="gen_docx"):
            with st.spinner("Generating DOCX..."):
                resp = requests.post(
                    f"{BASE_URL}/api/generate/docx",
                    json={
                        "job_id": st.session_state["job_id"],
                        "data":   st.session_state["data"],
                    },
                    timeout=60,
                )
            if resp.status_code == 200:
                st.session_state["docx_bytes"] = resp.content
                st.success("✅ Ready!")
            else:
                st.error(f"Failed: {resp.text}")

        if "docx_bytes" in st.session_state:
            st.download_button(
                label="⬇️ Download DOCX",
                data=st.session_state["docx_bytes"],
                file_name=f"Commissioni_{st.session_state['job_id']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary",
            )

    # ── Excel ──
    with col_e:
        st.subheader("📊 Excel Spreadsheet")
        st.caption("Excel with summary sheet + data")
        if st.button("Generate Excel", use_container_width=True, key="gen_excel"):
            with st.spinner("Generating Excel..."):
                resp = requests.post(
                    f"{BASE_URL}/api/generate/excel",
                    json={
                        "job_id": st.session_state["job_id"],
                        "data":   st.session_state["data"],
                    },
                    timeout=60,
                )
            if resp.status_code == 200:
                st.session_state["excel_bytes"] = resp.content
                st.success("✅ Ready!")
            else:
                st.error(f"Failed: {resp.text}")

        if "excel_bytes" in st.session_state:
            st.download_button(
                label="⬇️ Download Excel",
                data=st.session_state["excel_bytes"],
                file_name=f"Commissioni_{st.session_state['job_id']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

    # ── Both ──
    with col_b:
        st.subheader("📦 Both Files")
        st.caption("Generate DOCX + Excel together")
        if st.button("Generate Both", use_container_width=True, key="gen_both"):
            with st.spinner("Generating both files..."):
                resp = requests.post(
                    f"{BASE_URL}/api/generate/both",
                    json={
                        "job_id": st.session_state["job_id"],
                        "data":   st.session_state["data"],
                    },
                    timeout=120,
                )
            if resp.status_code == 200:
                result = resp.json()
                if result.get("docx_ready"):
                    st.success("✅ DOCX ready")
                if result.get("excel_ready"):
                    st.success("✅ Excel ready")
                # Fetch both files
                for fmt, key, fname, mime in [
                    ("docx",  "docx_bytes",  "docx",  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                    ("excel", "excel_bytes", "xlsx",  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                ]:
                    dl = requests.get(
                        f"{BASE_URL}/api/download/{st.session_state['job_id']}/{fmt}",
                        timeout=30,
                    )
                    if dl.status_code == 200:
                        st.session_state[key] = dl.content
            else:
                st.error(f"Failed: {resp.text}")

st.divider()

# ── Debug ────────────────────────────────────────────────────────────────────
if "job_id" in st.session_state:
    with st.expander("🛠️ Debug — Raw Extraction Details", expanded=False):
        resp = requests.post(
            f"{BASE_URL}/api/debug",
            json={"job_id": st.session_state["job_id"]},
            timeout=15,
        )
        if resp.status_code == 200:
            st.json(resp.json())
        else:
            st.error(resp.text)

st.caption("MetalloDoc v3.0 · Offline · No data leaves this machine")























#-----clientine kaanicha code using claude ai---------------------------- 
# import streamlit as st
# import requests
# import pandas as pd

# BASE_URL = "http://localhost:8000"


# st.set_page_config(page_title="PDF → DOCX Extractor", layout="wide")

# st.title("📄 PDF → DOCX Contract Extractor (Claude Test)")

# # ─────────────────────────────
# # Upload PDF
# # ─────────────────────────────
# st.header("📤 Upload PDF")

# uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

# if st.button("🚀 Analyze PDF"):
#     if uploaded_file:
#         with st.spinner("Processing with AI..."):

#             files = {
#                 "file": (uploaded_file.name, uploaded_file.getvalue(), "application/pdf")
#             }

#             response = requests.post(f"{BASE_URL}/api/upload", files=files)

#         if response.status_code == 200:
#             data = response.json()

#             st.success("✅ Extraction Done!")

#             st.session_state["job_id"] = data["job_id"]
#             st.session_state["data"] = data["data"]

#             st.subheader("📊 Extracted Data")
#             st.json(data)

#         else:
#             st.error(response.text)
#     else:
#         st.warning("Please upload a PDF")

# # ─────────────────────────────
# # Preview Table
# # ─────────────────────────────
# st.header("📋 Preview Data")

# if "job_id" in st.session_state:
#     if st.button("🔍 Load Preview"):
#         response = requests.post(
#             f"{BASE_URL}/api/preview",
#             json={"job_id": st.session_state["job_id"]}
#         )

#         if response.status_code == 200:
#             preview_data = response.json()

#             contracts = preview_data.get("contracts", [])

# # In app.py (Streamlit), replace the preview section:
#             for i, contract in enumerate(contracts):
#                 st.subheader(f"Contract {i+1}")
                
#                 rows = []
#                 for row in contract["payment_schedule"]:
#                     rows.append({
#                         "Mese": row.get("mese", ""),
#                         "Data Pagamento Commissioni": row.get("data_pagamento_commissioni", ""),
#                         "Management Mensile": row.get("management_mensile", ""),
#                     })
                
#                 df = pd.DataFrame(rows)
#                 st.dataframe(df, use_container_width=True)



# # ─────────────────────────────
# # Generate DOCX
# # ─────────────────────────────
# st.header("⬇️ Generate DOCX")

# if "job_id" in st.session_state:
#     if st.button("📄 Generate Document"):
#         response = requests.post(
#             f"{BASE_URL}/api/generate",
#             json={
#                 "job_id": st.session_state["job_id"],
#                 "data": st.session_state["data"]
#             }
#         )

#         if response.status_code == 200:
#             st.download_button(
#                 label="📥 Download DOCX",
#                 data=response.content,
#                 file_name="contract_output.docx",
#                 mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#             )
#         else:
#             st.error("Generation failed")


















# import streamlit as st
# import requests
# import pandas as pd

# BASE_URL = "http://localhost:8000"

# st.set_page_config(page_title="PDF → DOCX Extractor", layout="wide")

# st.title("📄 PDF → DOCX Contract Extractor (Claude Test)")

# # ─────────────────────────────
# # Upload PDF
# # ─────────────────────────────
# st.header("📤 Upload PDF")

# uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

# if st.button("🚀 Analyze PDF"):
#     if uploaded_file:
#         with st.spinner("Processing with AI..."):

#             files = {
#                 "file": (uploaded_file.name, uploaded_file.getvalue(), "application/pdf")
#             }

#             response = requests.post(f"{BASE_URL}/api/upload", files=files)

#         if response.status_code == 200:
#             data = response.json()

#             st.success("✅ Extraction Done!")

#             st.session_state["job_id"] = data["job_id"]
#             st.session_state["data"] = data["data"]

#             st.subheader("📊 Extracted Data")
#             st.json(data)

#         else:
#             st.error(response.text)
#     else:
#         st.warning("Please upload a PDF")

# # ─────────────────────────────
# # Preview Table
# # ─────────────────────────────
# st.header("📋 Preview Data")

# if "job_id" in st.session_state:
#     if st.button("🔍 Load Preview"):
#         response = requests.post(
#             f"{BASE_URL}/api/preview",
#             json={"job_id": st.session_state["job_id"]}
#         )

#         if response.status_code == 200:
#             preview_data = response.json()

#             contracts = preview_data.get("contracts", [])

#             for i, contract in enumerate(contracts):
#                 st.subheader(f"Contract {i+1}")

#                 df = pd.DataFrame(contract["payment_schedule"])
#                 st.dataframe(df, use_container_width=True)
#         else:
#             st.error("Preview failed")

# # ─────────────────────────────
# # Generate DOCX
# # ─────────────────────────────
# st.header("⬇️ Generate DOCX")

# if "job_id" in st.session_state:
#     if st.button("📄 Generate Document"):
#         response = requests.post(
#             f"{BASE_URL}/api/generate",
#             json={
#                 "job_id": st.session_state["job_id"],
#                 "data": st.session_state["data"]
#             }
#         )

#         if response.status_code == 200:
#             st.download_button(
#                 label="📥 Download DOCX",
#                 data=response.content,
#                 file_name="contract_output.docx",
#                 mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#             )
#         else:
#             st.error("Generation failed")