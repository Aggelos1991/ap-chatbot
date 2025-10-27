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
st.set_page_config(page_title="Vendor Statement Extractor", layout="wide")
st.title("Vendor Statement → Excel (Spanish • English • Greek)")

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
uploaded = st.file_uploader("Upload Vendor Statement (PDF)", type="pdf")

if uploaded:
    with st.spinner("Extracting text from PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded))

    st.text_area("Text Preview", text[:2000], height=150, disabled=True)

    if st.button("Extract to Excel", type="primary"):
        with st.spinner("Analyzing with GPT..."):
            try:
                data = extract_with_llm(text)
            except Exception as e:
                st.error(f"LLM failed: {e}")
                st.stop()

        # =============================================
        # FINAL LOGIC: PAYMENTS → CREDIT | RETENCIONES → SKIP
        # =============================================
        PAYMENT_TRIGGERS = [
            # Spanish
            r"\bcobro\b", r"\bpago\b", r"\babono\b", r"\bingreso\b", r"\brecibido\b",
            r"\bpago\s+recibido\b", r"\bcn\b", r"\breason\b",
            # English
            r"\bpayment\b", r"\breceipt\b", r"\breceived\b", r"\bcredit\b", r"\bcredited\b",
            # Greek
            r"\bπληρωμ[ήη]\b", r"\bπληρωμη\b", r"\bείσπραξ[ηη]\b", r"\bεισπραξη\b",
            r"\bκατάθεσ[ηη]\b", r"\bκαταθεση\b", r"\bπίστωσ[ηη]\b", r"\bπιστωση\b"
        ]

        IGNORE_TRIGGERS = [
            r"\bretenci[óo]n\b", r"\bwithholding\b",
            r"\bπαρακράτησ[ηη]\b", r"\bπαρακρατηση\b"
        ]

        cleaned_data = []
        for row in data:
            desc = " " + str(row.get("Description", "")).lower() + " "

            # 1. SKIP retenciones / withholdings
            if any(re.search(p, desc, re.IGNORECASE) for p in IGNORE_TRIGGERS):
                continue

            # 2. DETECT PAYMENT → FORCE TO CREDIT
            if any(re.search(p, desc, re.IGNORECASE) for p in PAYMENT_TRIGGERS):
                credit_val = float(row.get("Debit", 0) or row.get("Credit", 0))
                row["Credit"] = credit_val
                row["Debit"] = 0
            else:
                row["Debit"] = float(row.get("Debit", 0))
                row["Credit"] = float(row.get("Credit", 0))

            cleaned_data.append(row)

        data = cleaned_data
        # =============================================

        # 1. Show Raw JSON
        st.subheader("Raw JSON from LLM")
        st.json(data, expanded=False)

        # 2. Show Clean Table
        df = pd.DataFrame(data)
        for col in ["Debit", "Credit", "Balance"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        st.subheader("Final Extracted Table")
        st.dataframe(df, use_container_width=True)

        # 3. Download Excel
        excel_data = to_excel_bytes(data)
        st.download_button(
            label="Download Excel File",
            data=excel_data,
            file_name="vendor_statement_clean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Done! Payments → Credit | Retenciones → Removed")
