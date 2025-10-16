import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Hybrid Vendor Statement Extractor (Invoice / Payment / Credit Note)")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No OpenAI API key found. Add it to .env or Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(value):
    """Normalize decimals like 1.234,56 → 1234.56"""
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

def extract_raw_lines(uploaded_pdf):
    """Extract text lines from all pages of the PDF."""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", line):
                    all_lines.append(" ".join(line.split()))
    return all_lines

# ==========================================================
# GPT EXTRACTOR (Invoice / Payment / Credit Note)
# ==========================================================
def extract_with_gpt(lines):
    """Extract invoice, payment, and credit note lines, one amount column (Debit/Credit combined)."""
    BATCH_SIZE = 150
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a multilingual accountant specialized in Spanish and Greek vendor statements.

Below are text lines from a vendor statement. Some include:
- Spanish: "Fra. emitida", "Factura", "Abono", "Nota de Credito", "Cobro", "Pago", "Remesa", "Efecto"
- Greek: "Τιμολόγιο", "Πληρωμή", "Πιστωτικό", "Ακυρωτικό", "Τραπεζικό Έμβασμα"

Each line may contain one or more amounts (like 322,27 or 1.457,65).

Your job:
Extract only valid accounting lines, and for each line return:
- "Alternative Document": the document number (after Nº, n°, n., Factura, Documento, or similar)
- "Date": dd/mm/yy or dd/mm/yyyy if visible
- "Reason": one of these → "Invoice", "Payment", or "Credit Note"
- "Amount": the main numeric value (use the **last numeric value** in the line if unsure)

Rules:
- If the line contains "Abono", "Nota de Credito", "NC", "πιστω", "ακυρωτικ" → Reason = "Credit Note"
- If the line contains "Pago", "Cobro", "Remesa", "Efecto", "Transferencia", "Πληρωμή", "Τράπεζα", "Έμβασμα", "Μεταφορά" → Reason = "Payment"
- Otherwise → Reason = "Invoice"
- Ignore "Saldo", "Apertura", "Total General", "Base", "IVA", "FPA", "Υπόλοιπο", etc.
Output must be a valid JSON array.

Lines:
\"\"\"{text_block}\"\"\"
"""

        try:
            response = client.responses.create(model=MODEL, input=prompt)
            content = response.output_text.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            if not json_match:
                continue
            data = json.loads(json_match.group(0))
        except Exception as e:
            st.warning(f"⚠️ GPT failed on batch {i//BATCH_SIZE + 1}: {e}")
            continue

        for row in data:
            amount = normalize_number(row.get("Amount"))
            reason = str(row.get("Reason", "")).lower()

            # Sign logic: payments and credit notes become negative
            if any(k in reason for k in ["payment", "cobro", "pago", "remesa", "efecto", "transferencia", "πληρωμή", "τραπεζ", "έμβασμα", "μεταφορά"]):
                amount = -abs(amount)
            elif any(k in reason for k in ["credit", "abono", "nota de credito", "nc", "πιστω", "ακυρωτικ"]):
                amount = -abs(amount)

            all_records.append({
                "Alternative Document": str(row.get("Alternative Document", "")).strip(),
                "Date": str(row.get("Date", "")).strip(),
                "Reason": row.get("Reason", "").strip(),
                "Amount": amount
            })

    return all_records

# ==========================================================
# EXPORT
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from all pages..."):
        lines = extract_raw_lines(uploaded_pdf)

    if not lines:
        st.warning("⚠️ No readable text lines found. Check if the PDF is scanned.")
    else:
        st.text_area("📄 Preview (first 25 lines):", "\n".join(lines[:25]), height=250)

        if st.button("🤖 Run Hybrid Extraction"):
            with st.spinner("Analyzing data with GPT-4o-mini..."):
                data = extract_with_gpt(lines)

            if not data:
                st.warning("⚠️ No structured invoice data detected.")
            else:
                df = pd.DataFrame(data)
                st.success(f"✅ Extraction complete — {len(df)} valid records found.")
                st.dataframe(df, use_container_width=True)
                st.download_button(
                    "⬇️ Download Excel",
                    data=to_excel_bytes(data),
                    file_name="vendor_statement_invoices_payments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("Please upload a vendor statement PDF to begin.")
