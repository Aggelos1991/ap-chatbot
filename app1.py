import os, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# =============================================
# STREAMLIT CONFIG
# =============================================
st.set_page_config(page_title="🦅 DataFalcon — DEBE Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor (DEBE Calibrated)")

# =============================================
# HELPERS
# =============================================
def extract_text_from_pdf(file):
    """Safely extract text from uploaded PDF."""
    file_bytes = file.getvalue()
    if not file_bytes:
        raise ValueError("Uploaded file is empty or unreadable.")

    text = ""
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text


def clean_text(text):
    """Clean extracted text."""
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())


def normalize_number(value):
    """Normalize European-style numbers."""
    if not value:
        return ""
    s = str(value).strip()
    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")
    else:
        s = re.sub(r"[^\d.]", "", s)
    return s


def extract_tax_id(raw_text):
    """Detect CIF/NIF/VAT from text."""
    patterns = [
        r"\b[A-Z]{1}\d{7}[A-Z0-9]{1}\b",
        r"\bES\d{9}\b",
        r"\bEL\d{9}\b",
        r"\b[A-Z]{2}\d{8,12}\b",
    ]
    for pat in patterns:
        match = re.search(pat, raw_text)
        if match:
            return match.group(0)
    return None

# =============================================
# CORE EXTRACTION (DEBE-FIRST LOGIC)
# =============================================
def extract_debe_lines(raw_text):
    """
    Extract only invoice lines where 'Fra. emitida' and invoice pattern (6--) exist.
    Takes the first numeric amount (DEBE) and ignores the rest.
    """

    pattern = re.compile(
        r"(?P<debe>\d{1,3}(?:[\.,]\d{2,3})+)\s+\d{1,2}/\d{1,2}/\d{2,4}.*?(?P<doc>6[-–]\d{1,5}).*?(?P<date>\d{1,2}/\d{1,2}/\d{2,4}).*?(Fra\. emitida|Factura|Doc)",
        re.IGNORECASE,
    )

    rows = []
    for match in pattern.finditer(raw_text):
        val = normalize_number(match.group("debe"))
        date = match.group("date")
        doc = match.group("doc")

        try:
            amount = float(val)
        except:
            continue

        if amount <= 0 or amount > 100000:
            continue

        rows.append({
            "Alternative Document": doc.strip(),
            "Date": date.strip(),
            "Reason": "Invoice",
            "Document Value": f"{amount:.2f}"
        })

    df = pd.DataFrame(rows).drop_duplicates(subset=["Alternative Document", "Date"])
    return df


def to_excel_bytes(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# =============================================
# STREAMLIT UI
# =============================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded_pdf))

    st.text_area("🔍 Extracted Text Preview", text[:2500], height=250)

    if st.button("🔎 Extract DEBE Invoices"):
        df = extract_debe_lines(text)
        if not df.empty:
            tax_id = extract_tax_id(text)
            df["Tax ID"] = tax_id if tax_id else "Missing TAX ID"
            st.success(f"✅ Extracted {len(df)} invoices (DEBE only, SALDO ignored)")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "⬇️ Download Excel (Vendor Statement)",
                data=to_excel_bytes(df),
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No DEBE-based invoices found. Try uploading another page or format.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
