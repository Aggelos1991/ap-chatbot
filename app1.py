import os, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# =============================================
# CONFIG
# =============================================
st.set_page_config(page_title="🦅 DataFalcon — DEBE Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor (Accurate DEBE Version)")

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


def clean_text(text: str) -> str:
    """Normalize whitespace and remove line breaks."""
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())


def normalize_number(value):
    """Normalize EU formatted numbers like 1.234,56 → 1234.56."""
    if not value:
        return ""
    s = str(value).strip()
    s = re.sub(r"[^\d,\.]", "", s)
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
# CORE EXTRACTION LOGIC
# =============================================
def extract_debe_lines(raw_text):
    """
    Extracts invoice lines (Fra. emitida nº / Factura) with DEBE values only.
    Logic:
      - First numeric group before next HABER/SALDO is DEBE.
      - Ignore lines containing Cobro, Banco, Apertura, Pago, Saldo.
    """

    # Split the text into pseudo-lines (every 'Fra. emitida' marks a new invoice)
    fragments = re.split(r"(?i)(?=Fra\. emitida|Factura\s?n[ºo]?|SerieFactura)", raw_text)
    results = []

    for frag in fragments:
        frag = frag.strip()
        if not frag:
            continue
        # Skip irrelevant lines
        if any(word in frag.lower() for word in ["cobro", "banco", "pago", "saldo", "apertura", "efec"]):
            continue

        # Document number
        doc_match = re.search(r"\b\d{1,2}\s*[-–]{1,2}\s*\d{3,5}\b|\b6[-–]\d{3,5}\b", frag)
        doc = doc_match.group(0).replace(" ", "") if doc_match else ""

        # Date (dd/mm/yy)
        date_match = re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", frag)
        date = date_match.group(0) if date_match else ""

        # Find numeric groups (DEBE, HABER, SALDO)
        nums = re.findall(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", frag)
        if not nums:
            continue

        # The DEBE is almost always the first valid number after Fra. emitida or near date
        debe_val = None
        for n in nums:
            val = normalize_number(n)
            if val:
                try:
                    f = float(val)
                    if 0 < f < 100000:  # realistic range
                        debe_val = f
                        break
                except:
                    continue

        if not debe_val:
            continue

        results.append({
            "Alternative Document": doc,
            "Date": date,
            "Reason": "Invoice",
            "Document Value": f"{debe_val:.2f}"
        })

    df = pd.DataFrame(results).drop_duplicates(subset=["Alternative Document", "Date"])
    return df

# =============================================
# EXPORT
# =============================================
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

    if st.button("🧾 Extract DEBE Values"):
        df = extract_debe_lines(text)

        if not df.empty:
            tax_id = extract_tax_id(text)
            df["Tax ID"] = tax_id if tax_id else "Missing TAX ID"
            st.success(f"✅ Extracted {len(df)} invoices (DEBE only, Haber/Saldo ignored).")
            st.dataframe(df, use_container_width=True)
            st.download_button(
                "⬇️ Download Excel (Vendor Statement)",
                data=to_excel_bytes(df),
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No DEBE lines detected. Try another PDF or share sample text.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
