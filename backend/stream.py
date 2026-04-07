import streamlit as st
import requests
import pandas as pd

BASE_URL = "http://localhost:8000"

st.set_page_config(page_title="PDF → DOCX Extractor", layout="wide")

st.title("📄 PDF → DOCX Contract Extractor (Claude Test)")

# ─────────────────────────────
# Upload PDF
# ─────────────────────────────
st.header("📤 Upload PDF")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if st.button("🚀 Analyze PDF"):
    if uploaded_file:
        with st.spinner("Processing with AI..."):

            files = {
                "file": (uploaded_file.name, uploaded_file.getvalue(), "application/pdf")
            }

            response = requests.post(f"{BASE_URL}/api/upload", files=files)

        if response.status_code == 200:
            data = response.json()

            st.success("✅ Extraction Done!")

            st.session_state["job_id"] = data["job_id"]
            st.session_state["data"] = data["data"]

            st.subheader("📊 Extracted Data")
            st.json(data)

        else:
            st.error(response.text)
    else:
        st.warning("Please upload a PDF")

# ─────────────────────────────
# Preview Table
# ─────────────────────────────
st.header("📋 Preview Data")

if "job_id" in st.session_state:
    if st.button("🔍 Load Preview"):
        response = requests.post(
            f"{BASE_URL}/api/preview",
            json={"job_id": st.session_state["job_id"]}
        )

        if response.status_code == 200:
            preview_data = response.json()

            contracts = preview_data.get("contracts", [])

# In app.py (Streamlit), replace the preview section:
            for i, contract in enumerate(contracts):
                st.subheader(f"Contract {i+1}")
                
                rows = []
                for row in contract["payment_schedule"]:
                    rows.append({
                        "Mese": row.get("mese", ""),
                        "Data Pagamento Commissioni": row.get("data_pagamento_commissioni", ""),
                        "Management Mensile": row.get("management_mensile", ""),
                    })
                
                df = pd.DataFrame(rows)
                st.dataframe(df, use_container_width=True)



# ─────────────────────────────
# Generate DOCX
# ─────────────────────────────
st.header("⬇️ Generate DOCX")

if "job_id" in st.session_state:
    if st.button("📄 Generate Document"):
        response = requests.post(
            f"{BASE_URL}/api/generate",
            json={
                "job_id": st.session_state["job_id"],
                "data": st.session_state["data"]
            }
        )

        if response.status_code == 200:
            st.download_button(
                label="📥 Download DOCX",
                data=response.content,
                file_name="contract_output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("Generation failed")


















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