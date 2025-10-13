import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
import re

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Accurate DEBE Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Vendor Statement Extractor (Structured PDF Mode)")

# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(value):
    """Normalize decimals: 1.234,56 → 1234.56"""
    if not value:
        return ""
    s = str(value).strip().replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return round(float(s), 2)
    except:
        return ""

def extract_table_data(uploaded_pdf):
    """Extract DEBE column data from structured table layout."""
    records = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue

            headers = [h.strip().upper() if h else "" for h in table[0]]
            for row in table[1:]:
                if not any(row):
                    continue

                # try to find DEBE, HABER, SALDO positions
                try:
                    debe_idx = next(i for i, h in enumerate(headers) if "DEBE" in h)
                    haber_idx = next(i for i, h in enumerate(headers) if "HABER" in h)
                    saldo_idx = next(i for i, h in enumerate(headers) if "SALDO" in h)
                except StopIteration:
                    continue

                # extract core columns
                date = next((x for x in row if re.match(r"\d{2}/\d{2}/\d{2,4}", str(x))), "")
                doc = row[3] if len(row) > 3 else ""
                debe = row[debe_idx] if len(row) > debe_idx else ""
                haber = row[haber_idx] if len(row) > haber_idx else ""
                saldo = row[saldo_idx] if len(row) > saldo_idx else ""

                # ignore zeros and empty DEBE
                if not debe or re.match(r"^0+[.,]0+$", str(debe)):
                    continue

                # detect credit notes
                concept = " ".join(str(x) for x in row if x)
                reason = "Credit Note" if re.search(r"(?i)(ABONO|CREDIT|NOTA\s+DE\s+CR[EÉ]DITO)", concept) else "Invoice"

                records.append({
                    "Alternative Document": doc,
                    "Date": date,
                    "Reason": reason,
                    "Document Value": normalize_number(debe),
                })
    return records

def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT APP
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Reading structured table data..."):
        try:
            data = extract_table_data(uploaded_pdf)
        except Exception as e:
            st.error(f"❌ PDF extraction failed: {e}")
            st.stop()

    if not data:
        st.warning("⚠️ No DEBE records detected. Check if the PDF is tabular or scanned.")
    else:
        df = pd.DataFrame(data)
        st.success(f"✅ Extraction complete — {len(df)} valid DEBE entries found.")
        st.dataframe(df, use_container_width=True)
        st.download_button(
            "⬇️ Download Excel",
            data=to_excel_bytes(data),
            file_name="vendor_statement_structured.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Please upload a vendor statement PDF to begin.")
