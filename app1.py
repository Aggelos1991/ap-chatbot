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
    """Analyze extracted lines using GPT-4o-mini for invoices, credit notes, and payment detections (DEBE + HABER)."""
    BATCH_SIZE = 200
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a multilingual accountant specialized in Spanish and Greek vendor statements.

Each line may include:
- Spanish: DEBE (debit), HABER (credit), TOTAL, SALDO, COBRO, EFECTO, REMESA
- Greek: ΧΡΕΩΣΗ (debit), ΠΙΣΤΩΣΗ (credit), ΣΥΝΟΛΟ, ΠΛΗΡΩΜΗ, ΤΡΑΠΕΖΑ, ΕΜΒΑΣΜΑ

Your task:
For each valid accounting line, extract:
- "Alternative Document": document number (Documento, Factura, Τιμολόγιο, Παραστατικό, etc.)
- "Date": dd/mm/yy or dd/mm/yyyy
- "Reason": short description (e.g., "Factura", "Abono", "Πληρωμή", "Τραπεζικό Έμβασμα")
- "DEBE Value": numeric amount under DEBE or ΧΡΕΩΣΗ
- "HABER Value": numeric amount under HABER, ΠΙΣΤΩΣΗ, COBRO, or similar

Rules:
1. If both DEBE and HABER (or ΧΡΕΩΣΗ and ΠΙΣΤΩΣΗ) appear:
   - Assign DEBE → "DEBE Value" (Debit)
   - Assign HABER → "HABER Value" (Credit)
   - Ignore TOTAL or ΣΥΝΟΛΟ in this case.
2. Use TOTAL or ΣΥΝΟΛΟ only if DEBE/HABER are absent.
3. If the text contains "Abono", "Nota de Crédito", "Πιστωτικό", or "Ακυρωτικό" → classify as Credit Note.
4. If it contains "Pago", "Cobro", "Remesa", "Efecto", "Πληρωμή", "Τράπεζα", "Έμβασμα", "Μεταφορά" → classify as Payment.
5. If neither → classify as Invoice.
6. Ignore summary lines (Saldo, IVA, Impuesto, Υπόλοιπο, ΦΠΑ, Βάση, Υποσύνολο, etc.)
7. Output a valid JSON array with numeric strings (use '.' for decimals).

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
            # Normalize and parse values
            debe_val = normalize_number(row.get("DEBE Value"))
            haber_val = normalize_number(row.get("HABER Value"))
            val = normalize_number(row.get("Document Value")) or debe_val
            pay = haber_val or normalize_number(row.get("Payment Value"))

            reason_text = str(row.get("Reason", "")).lower()

            # --- safety fallback: if "haber"/"πίστωση" appears but GPT filled only val ---
            if (("haber" in reason_text or "πίστ" in reason_text) and val and not pay):
                pay, val = val, 0.0

            # --- classify by reason ---
            if any(k in reason_text for k in ["abono", "credit", "nota de crédito", "nc", "πιστω", "ακυρωτικ"]):
                val = -abs(val)
                doc_type = "Credit Note"
            elif any(k in reason_text for k in ["pago", "remesa", "cobro", "efecto", "transferencia", "πληρωμή", "τράπεζ", "έμβασμα", "μεταφορά"]):
                doc_type = "Payment"
            else:
                doc_type = "Invoice"

            all_records.append({
                "Alternative Document": str(row.get("Alternative Document", "")).strip(),
                "Date": str(row.get("Date", "")).strip(),
                "Reason": doc_type,
                "Document Value": val,
                "Payment Value": pay
            })

    return all_records

# ==========================================================
# EXCEL EXPORT
# ==========================================================
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
