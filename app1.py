import os, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# =============================================
# STREAMLIT CONFIG
# =============================================
st.set_page_config(page_title="🦅 DataFalcon — Accurate DEBE Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor (DEBE Accurate Edition)")

# =============================================
# HELPERS
# =============================================
def extract_text_from_pdf(file):
    """Extract text from PDF safely."""
    file_bytes = file.getvalue()
    if not file_bytes:
        raise ValueError("Uploaded file is empty.")
    text = ""
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text


def normalize_number(value):
    """Convert EU-format numbers like 1.234,56 → 1234.56"""
    s = str(value).strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return round(float(s), 2)
    except:
        return None


def extract_tax_id(text):
    """Extract CIF/NIF/VAT"""
    match = re.search(r"\b([A-Z]{1}\d{7}[A-Z0-9]{1}|ES\d{9}|EL\d{9}|[A-Z]{2}\d{8,12})\b", text)
    return match.group(0) if match else "Missing TAX ID"


def extract_invoice_lines(text):
    """
    Extract invoice data from PDFs with or without line breaks.
    Logic:
      - Look for '6--xxx' or 'Fra. emitida nº:'
      - Capture the nearest DEBE number (first numeric before next number)
      - Ignore lines with SALDO, HABER, BANCO, COBRO, EFECTO
    """
    records = []

    # Pre-clean
    txt = text.replace("\n", " ").replace("\xa0", " ")
    txt = re.sub(r"\s+", " ", txt)

    # Find all potential invoice groups
    pattern = re.compile(
        r"(Fra\. emitida nº:\s*|Factura|SerieFactura.*?Num FactCliente\s*|Doc\s*:?|6--\d+)"
        r".{0,80}?(\d{1,3}(?:[\.,]\d{3})*[\.,]\d{2})"   # first numeric value = DEBE
        r".{0,20}?(\d{2}/\d{2}/\d{2,4})",              # date
        re.IGNORECASE
    )

    for match in re.finditer(pattern, txt):
        chunk = match.group(0)

        # Skip payments or balances
        if any(skip in chunk.upper() for skip in ["SALDO", "HABER", "BANCO", "COBRO", "EFECTO"]):
            continue

        inv_match = re.search(r"(\d+--\d+|\d{1,4}/\d{1,4})", chunk)
        inv_no = inv_match.group(1) if inv_match else ""

        date_match = re.search(r"\b\d{2}/\d{2}/\d{2,4}\b", chunk)
        date_val = date_match.group(0) if date_match else ""

        nums = re.findall(r"\d{1,3}(?:[\.,]\d{3})*[\.,]\d{2}", chunk)
        nums = [normalize_number(n) for n in nums if n not in ["0,00", "0.00"] and normalize_number(n)]
        if not nums:
            continue

        doc_value = nums[0]
        reason = "Credit Note" if any(k in chunk.lower() for k in ["abono", "nota de crédito", "nc", "credit"]) else "Invoice"

        records.append({
            "Alternative Document": inv_no,
            "Date": date_val,
            "Reason": reason,
            "Document Value": doc_value
        })

    return records


def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# =============================================
# STREAMLIT APP
# =============================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        try:
            text = extract_text_from_pdf(uploaded_pdf)
        except Exception as e:
            st.error(f"❌ Failed to read PDF: {e}")
            st.stop()

    st.text_area("🔍 Extracted Text Preview", text[:2000], height=200)

    tax_id = extract_tax_id(text)
    data = extract_invoice_lines(text)

    if not data:
        st.warning("⚠️ No valid invoice data found. Try another PDF or verify text layout.")
    else:
        for row in data:
            row["Tax ID"] = tax_id

        df = pd.DataFrame(data)
        st.success(f"✅ Extraction complete — {len(df)} invoices found (DEBE only, Saldo/Haber ignored).")
        st.dataframe(df, use_container_width=True)

        st.download_button(
            "⬇️ Download Excel",
            data=to_excel_bytes(data),
            file_name="vendor_statement_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Please upload a vendor statement PDF to begin.")
