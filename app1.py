import os
import json
import re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# ==========================
# CONFIG
# ==========================
st.set_page_config(page_title="ABONO = CREDIT", layout="wide")
st.title("Vendor Statement → Excel (ABONO = CREDIT)")

API_KEY = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not API_KEY:
    st.error("OpenAI API key missing.")
    st.stop()

client = OpenAI(api_key=API_KEY)
MODEL = "gpt-4o-mini"

# ==========================
# HELPERS
# ==========================
def extract_text_from_pdf(file):
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())

def extract_with_llm(raw_text):
    prompt = f"""
    Extract every invoice line from the vendor statement.
    Return **ONLY** a JSON array of objects with these exact keys:
    - Invoice_Number (string)
    - Date (string, DD/MM/YYYY or YYYY-MM-DD)
    - Description (string)
    - Debit (number, 0 if empty)
    - Credit (number, 0 if empty)
    - Balance (number, 0 if empty)

    Text (max 12,000 chars):
    \"\"\"{raw_text[:12000]}\"\"\"
    """
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Return ONLY valid JSON. No markdown, no explanations."},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        max_tokens=1500
    )
    json_str = resp.choices[0].message.content.strip()

    if "```" in json_str:
        parts = json_str.split("```")
        json_str = parts[1] if len(parts) > 1 else parts[0]
        if json_str.lower().startswith("json"):
            json_str = json_str[4:].strip()

    return json.loads(json_str)

def to_excel_bytes(records):
    output = BytesIO()
    pd.DataFrame(records).to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    return output

# ==========================
# UI
# ==========================
uploaded = st.file_uploader("Upload PDF", type="pdf")

if uploaded:
    with st.spinner("Reading PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded))

    st.text_area("Text Preview", text[:2000], height=150, disabled=True)

    if st.button("EXTRACT → ABONO = CREDIT", type="primary"):
        with st.spinner("GPT is working..."):
            data = extract_with_llm(text)

        # =============================================
        # ABONO = CREDIT. PAGO = CREDIT. NO EXCEPTIONS.
        # =============================================
        CREDIT_TRIGGERS = [
            "abono", "pago", "cobro", "transference", "transferencia",
            "ingreso", "recibido", "pago recibido", "cn", "nota de crédito",
            "credit note", "credit", "credited", "payment", "receipt",
            "πληρωμή", "πληρωμη", "είσπραξη", "εισπραξη", "κατάθεση",
            "μεταφορά", "πιστωτικό", "επιστροφή"
        ]

        IGNORE_TRIGGERS = [
            "retención", "retencion", "withholding",
            "παρακράτηση", "παρακρατηση"
        ]

        final_data = []
        for row in data:
            desc = str(row.get("Description", "")).lower()

            # 1. ABONO / PAGO / etc → FORCE CREDIT
            if any(trigger in desc for trigger in CREDIT_TRIGGERS):
                row["Credit"] = float(row.get("Debit", 0) or row.get("Credit", 0))
                row["Debit"] = 0
            else:
                row["Debit"] = float(row.get("Debit", 0))
                row["Credit"] = float(row.get("Credit", 0))

            # 2. DELETE RETENCIÓN
            if any(ignore in desc for ignore in IGNORE_TRIGGERS):
                continue

            final_data.append(row)

        data = final_data
        # =============================================

        # Show JSON
        st.subheader("Raw JSON")
        st.json(data, expanded=False)

        # Show Table
        df = pd.DataFrame(data)
        for col in ["Debit", "Credit", "Balance"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        st.subheader("ABONO = CREDIT → FINAL TABLE")
        st.dataframe(df, use_container_width=True)

        # Download
        st.download_button(
            label="DOWNLOAD EXCEL – ABONO IS CREDIT",
            data=to_excel_bytes(data),
            file_name="ABONO_IS_CREDIT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("ABONO = CREDIT. PAGO = CREDIT. DONE.")
