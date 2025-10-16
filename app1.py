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
st.title("🦅 DataFalcon Pro — Hybrid Vendor Statement Extractor (Optimized)")

# Load API key
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
    """Extract all text lines from every page of the PDF."""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for p_i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", line):
                    clean_line = " ".join(line.split())
                    all_lines.append(clean_line)
    return all_lines

# ==========================================================
# GPT EXTRACTOR
# ==========================================================
def extract_with_gpt(lines):
    """Analyze extracted lines using GPT-4o-mini for invoices, credit notes, and payment detections."""
    BATCH_SIZE = 200
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a multilingual accountant specializing in Spanish and Greek vendor statements.

Below are text lines from a vendor statement (possibly in Spanish, Greek, or English).

Each line may contain multiple numbers — usually labeled as:
- Spanish: DEBE, HABER, TOTAL, TOTALE, SALDO, COBRO, PAGO, EFECTO, REMESA
- Greek: ΧΡΕΩΣΗ, ΠΙΣΤΩΣΗ, ΣΥΝΟΛΟ, ΥΠΟΛΟΙΠΟ, ΠΛΗΡΩΜΗ, ΤΡΑΠΕΖΑ, ΤΡΑΠΕΖΙΚΟ ΕΜΒΑΣΜΑ, ΜΕΤΑΦΟΡΑ

Your job:
1. Extract only valid **invoice**, **credit note**, or **payment** lines.
2. For each line, return:
   - "Alternative Document": document number (under Documento, Num, Nº, Numero, N°, Factura, Τιμολόγιο, Παραστατικό, or similar)
   - "Date": dd/mm/yy or dd/mm/yyyy
   - "Reason": text describing the line (e.g. "Factura", "Abono", "Πληρωμή", "Τραπεζικό Έμβασμα")
   - "Document Value":
       • If the line contains DEBE, ΧΡΕΩΣΗ, TOTAL, take that numeric value.
       • If it's a Credit Note (ABONO, NOTA DE CRÉDITO, ΠΙΣΤΩΤΙΚΟ, ΑΚΥΡΩΤΙΚΟ), make the value negative.
   - "Payment Value":
       • If the line refers to payment or transfer (Cobro, Pago, Remesa, Efecto, Πληρωμή, Τράπεζα, Τραπεζικό Έμβασμα, Μεταφορά), 
         extract the HABER / ΠΙΣΤΩΣΗ / CREDIT numeric value.
       • If both DEBE and HABER appear, assign DEBE to "Document Value" and HABER to "Payment Value".
3. Ignore lines referring only to summaries:
   "Saldo anterior", "Total general", "Base", "IVA", "Impuesto", "Resumen", "Υπόλοιπο", "Προηγούμενο Υπόλοιπο", "ΦΠΑ", "Βάση", "Υποσύνολο".
4. Always output a valid JSON array.
5. Ensure numeric values use '.' for decimals and exactly two digits.
6. Do not leave empty or null document numbers when visible.

Lines:
\"\"\"{text_block}\"\"\"
"""

        try:
            response = client.responses.create(model=MODEL, input=prompt)
            content = response.output_text.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            json_text = json_match.group(0) if json_match else content
            data = json.loads(json_text)
        except Exception as e:
            st.warning(f"⚠️ GPT failed on batch {i//BATCH_SIZE + 1}: {e}")
            continue

        for row in data:
            val = normalize_number(row.get("Document Value"))
            pay = normalize_number(row.get("Payment Value"))

            if val == "" and pay == "":
                continue

            reason = row.get("Reason", "").lower()

            # --- Classify type ---
            if any(k in reason for k in ["abono", "credit", "nota de crédito", "nc", "πιστωτικό", "ακυρωτικό"]):
                val = -abs(val)
                doc_type = "Credit Note"
            elif any(k in reason for k in ["pago", "remesa", "cobro", "efecto", "transferencia", "πληρωμή", "τράπεζα", "έμβασμα", "μεταφορά"]):
                doc_type = "Payment"
            else:
                doc_type = "Invoice"

            all_records.append({
                "Alternative Document": row.get("Alternative Document", "").strip(),
                "Date": row.get("Date", "").strip(),
                "Reason": doc_type,
                "Document Value": val,
                "Payment Value": pay
            })

    return all_records

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
                    file_name="vendor_statement_hybrid.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("Please upload a vendor statement PDF to begin.") 
