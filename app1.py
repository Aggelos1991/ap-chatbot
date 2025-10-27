import os
import json
import re
from io import BytesIO
import fitz                     # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# ==========================
# CONFIG
# ==========================
st.set_page_config(page_title="DataFalcon Pro — Vendor Statement Parser", layout="wide")
st.markdown(
    """
    <h1 style='text-align:center; font-size:3rem; font-weight:700;
    background: linear-gradient(90deg,#0D47A1,#42A5F5);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent;'>DataFalcon Pro</h1>
    <h3 style='text-align:center; color:#1565C0;'>Vendor Statement to Excel (ABONO = CREDIT + REASON)</h3>
    """,
    unsafe_allow_html=True,
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
    """Extract raw text from uploaded PDF (PyMuPDF)."""
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text


def clean_text(text: str) -> str:
    """Normalize whitespace & currency symbols."""
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())


def extract_with_llm(raw_text: str):
    """Ask GPT for a clean JSON array of transaction lines."""
    prompt = f"""
    Extract **all** transaction lines from the vendor statement.
    Return **ONLY** a valid JSON array of objects with the keys:
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
            {"role": "system", "content": "Return ONLY valid JSON. No markdown, no explanations."},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        max_tokens=1500,
    )
    json_str = resp.choices[0].message.content.strip()

    # Strip possible markdown fences
    if "```" in json_str:
        json_str = re.search(r"```(?:json)?\s*(.*?)\s*```", json_str, re.S)
        json_str = json_str.group(1) if json_str else json_str

    try:
        return json.loads(json_str)
    except Exception:
        st.warning("GPT output not valid JSON – trying a quick sanitise.")
        json_str = re.sub(r"[^{}\[\]:,0-9A-Za-z.\-\"'/ ]", "", json_str)
        return json.loads(json_str)


def to_excel_bytes(records):
    """Return Excel file as bytes for download."""
    out = BytesIO()
    pd.DataFrame(records).to_excel(out, index=False, engine="openpyxl")
    out.seek(0)
    return out


# ==========================
# UI
# ==========================
uploaded = st.file_uploader("Upload Vendor Statement (PDF)", type="pdf")

if uploaded:
    with st.spinner("Reading PDF…"):
        raw_text = clean_text(extract_text_from_pdf(uploaded))

    st.text_area("Preview (first 2000 chars)", raw_text[:2000], height=150, disabled=True)

    with st.spinner("Extracting with GPT…"):
        data = extract_with_llm(raw_text)

    # ------------------------------------------------------------------
    #  SMART REASON / ABONO LOGIC (the version you had yesterday)
    # ------------------------------------------------------------------
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
    ABONO_TRIGGERS = ["abono", "abonos"]          # <-- the missing piece
    IGNORE_TRIGGERS = [
        "retención", "retencion", "withholding", "παρακράτηση", "παρακρατηση"
    ]

    final_rows = []
    raw_lower = raw_text.lower()

    for row in data:
        desc = str(row.get("Description", "")).lower()
        inv  = str(row.get("Invoice_Number", "")).lower()

        # ---- 1. Skip rows we never want ----
        if any(t in desc for t in IGNORE_TRIGGERS):
            continue

        debit  = float(row.get("Debit", 0) or 0)
        credit = float(row.get("Credit", 0) or 0)

        # ---- 2. Reason cascade (exact order you used yesterday) ----
        reason = "INVOICE"

        # a) Credit-Note
        if any(t in desc for t in CREDIT_NOTE_TRIGGERS):
            reason = "CREDIT NOTE"
            credit = credit if credit else debit
            debit  = 0

        # b) Payment
        elif any(t in desc for t in PAYMENT_TRIGGERS) or any(
            t in raw_lower and inv in raw_lower for t in PAYMENT_TRIGGERS
        ):
            reason = "PAYMENT"
            credit = credit if credit else debit
            debit  = 0

        # c) ABONO (fallback – any line that contains the word “abono”)
        elif any(t in desc for t in ABONO_TRIGGERS):
            reason = "ABONO"
            credit = credit if credit else debit
            debit  = 0

        # d) Default → regular invoice
        else:
            debit  = debit
            credit = credit

        # ---- 3. Assemble final row ----
        row.update({
            "Debit":  debit,
            "Credit": credit,
            "Reason": reason
        })
        final_rows.append(row)

    # ------------------------------------------------------------------
    #  DISPLAY & DOWNLOAD
    # ------------------------------------------------------------------
    df = pd.DataFrame(final_rows)

    # Drop any stray “Balance” column that sometimes appears
    if "Balance" in df.columns:
        df = df.drop(columns=["Balance"])

    # Force numeric
    for c in ["Debit", "Credit"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Reason first, then the rest
    cols = ["Reason"] + [c for c in df.columns if c != "Reason"]
    df = df[cols]

    st.subheader("Parsed Data (ABONO = CREDIT + REASON)")
    st.dataframe(df, use_container_width=True)

    st.download_button(
        label="DOWNLOAD EXCEL",
        data=to_excel_bytes(df.to_dict(orient="records")),
        file_name="DataFalcon_Pro_Statement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Extraction complete — every **ABONO** is now a **Credit** with a clear **Reason**.")
