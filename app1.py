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
st.set_page_config(page_title="🦅 DataFalcon Pro — Vendor Statement Parser", layout="wide")
st.markdown(
    """
    <h1 style='text-align:center; font-size:3rem; font-weight:700;
    background: linear-gradient(90deg,#0D47A1,#42A5F5);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent;'>🦅 DataFalcon Pro</h1>
    <h3 style='text-align:center; color:#1565C0;'>Vendor Statement → Excel (ABONO = CREDIT + REASON)</h3>
    """,
    unsafe_allow_html=True
)

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
    """Extract text from uploaded PDF using PyMuPDF"""
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    """Normalize text (remove nbsp, unify € symbol)"""
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())

def extract_with_llm(raw_text):
    """Ask GPT to extract invoice lines into JSON"""
    prompt = f"""
    Extract all transaction lines (invoices, payments, credits) from this vendor statement.
    Return ONLY a valid JSON array of objects with:
    - Invoice_Number (string)
    - Date (string, DD/MM/YYYY or YYYY-MM-DD)
    - Description (string)
    - Debit (number, 0 if empty)
    - Credit (number, 0 if empty)

    Text (max 12,000 chars):
    \"\"\"{raw_text[:12000]}\"\"\"
    """
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Return ONLY valid JSON. No markdown or commentary."},
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

    try:
        return json.loads(json_str)
    except Exception:
        st.warning("⚠️ GPT output not valid JSON, attempting auto-fix.")
        json_str = re.sub(r"[^{}\[\]:,0-9A-Za-z.\-\"'/ ]", "", json_str)
        return json.loads(json_str)

def to_excel_bytes(records):
    """Convert records to Excel bytes for download"""
    output = BytesIO()
    pd.DataFrame(records).to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return output

# ==========================
# UI
# ==========================
uploaded = st.file_uploader("📄 Upload Vendor Statement (PDF)", type="pdf")

if uploaded:
    with st.spinner("Reading PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded))

    st.text_area("Preview (first 2000 chars)", text[:2000], height=150, disabled=True)

    with st.spinner("Extracting data using GPT..."):
        data = extract_with_llm(text)

    # ==========================
    # SMART DETECTION LOGIC
    # ==========================
    PAYMENT_TRIGGERS = [
        "pago", "transferencia", "transference", "transfer", "remittance",
        "ingreso", "depósito", "deposito", "deposit", "bank", "payment",
        "paid", "pagado", "paid to", "bank transfer", "transferencia bancaria",
        "έμβασμα", "εμβασμα", "πληρωμή", "πληρωμη", "κατάθεση", "μεταφορά",
        "receipt", "received", "recibido", "transfered", "wire", "cash receipt"
    ]

    CREDIT_NOTE_TRIGGERS = [
        "nota de crédito", "nota credito", "credit note", "credit", "cn",
        "πιστωτικό", "πιστωτικο", "refund", "creditmemo", "credit memo", "credito"
    ]

    IGNORE_TRIGGERS = [
        "retención", "retencion", "withholding", "παρακράτηση", "παρακρατηση"
    ]

    final_data = []
    raw_lower = text.lower()

    for row in data:
        desc = str(row.get("Description", "")).lower()
        debit = float(row.get("Debit", 0) or 0)
        credit = float(row.get("Credit", 0) or 0)
        reason = "INVOICE"

        # Ignore retentions
        if any(ignore in desc for ignore in IGNORE_TRIGGERS):
            continue

        # --- Credit Notes ---
        if any(k in desc for k in CREDIT_NOTE_TRIGGERS):
            row["Credit"] = credit if credit != 0 else debit
            row["Debit"] = 0
            reason = "CREDIT NOTE"

        # --- Payments (detected in desc or in raw text if desc vague) ---
        elif any(k in desc for k in PAYMENT_TRIGGERS) or any(
            k in raw_lower for k in PAYMENT_TRIGGERS if str(row.get("Invoice_Number", "")).lower() in raw_lower
        ):
            row["Credit"] = credit if credit != 0 else debit
            row["Debit"] = 0
            reason = "PAYMENT"

        # --- Invoice Default ---
        else:
            row["Debit"] = debit
            row["Credit"] = credit
            reason = "INVOICE"

        row["Reason"] = reason
        final_data.append(row)

    # ==========================
    # CLEAN + DISPLAY
    # ==========================
    df = pd.DataFrame(final_data)
    # drop Balance if present
    if "Balance" in df.columns:
        df = df.drop(columns=["Balance"])

    for col in ["Debit", "Credit"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Reorder columns: Reason first
    cols = ["Reason"] + [c for c in df.columns if c != "Reason"]
    df = df[cols]

    st.subheader("🧾 Parsed Data (ABONO = CREDIT + REASON)")
    st.dataframe(df, use_container_width=True)

    st.download_button(
        label="📥 DOWNLOAD EXCEL",
        data=to_excel_bytes(df.to_dict(orient="records")),
        file_name="DataFalcon_Pro_Statement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ Extraction complete — ABONO = CREDIT. Reason column active (Invoice / Payment / Credit Note).")
