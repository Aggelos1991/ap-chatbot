import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# APP CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT+OCR Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Hybrid GPT + OCR Extractor")

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
PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"

# ==========================================================
# OCR-ENHANCED TEXT EXTRACTION
# ==========================================================
def extract_text_with_ocr(uploaded_pdf):
    """Extracts text from PDF using both pdfplumber and OCR fallback."""
    all_lines, ocr_pages = [], []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and len(text.strip()) > 10:
                for line in text.split("\n"):
                    clean_line = " ".join(line.split())
                    if clean_line:
                        all_lines.append(clean_line)
            else:
                # OCR fallback for scanned pages
                ocr_pages.append(i)
                img = convert_from_bytes(pdf_bytes, dpi=250, first_page=i, last_page=i)[0]
                ocr_text = pytesseract.image_to_string(img, lang="spa+eng+ell")
                for line in ocr_text.split("\n"):
                    clean_line = " ".join(line.split())
                    if clean_line:
                        all_lines.append(clean_line)

    return all_lines, ocr_pages

# ==========================================================
# GPT EXTRACTION
# ==========================================================
def normalize_number(value):
    """Normalize Spanish/Greek numeric formats like 1.234,56 → 1234.56"""
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

def parse_gpt_response(content, batch_num):
    json_match = re.search(r'\[.*\]', content, re.DOTALL)
    if not json_match:
        st.warning(f"⚠️ Batch {batch_num}: No JSON found. First 300 chars:\n{content[:300]}")
        return []
    try:
        return json.loads(json_match.group(0))
    except json.JSONDecodeError as e:
        st.warning(f"⚠️ Batch {batch_num}: JSON decode error → {e}")
        return []

def extract_with_gpt(lines):
    """GPT-based multilingual statement extraction with Spanish/Greek emphasis."""
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a multilingual financial data extractor (Spanish, Greek, English).

Analyze the following text lines from a vendor statement and extract structured data.

Possible column meanings:
- Date (Fecha / Ημερομηνία)
- Document or Reference (Documento / N° DOC / Αρ. Παραστατικού / Reference / Invoice)
- Description (Concepto / Περιγραφή / Description)
- DEBE / Χρέωση / Charge = Invoice amount
- HABER / Πίστωση / Credit = Payment or credit note amount
- TOTAL / TOTALES / ΣΥΝΟΛΟ / ΤΕΛΙΚΟ / IMPORTE TOTAL only used **if DEBE/HABER missing**
- SALDO lines should be ignored (running balance)
- “Referencia”, “Factura”, “Fra.”, “Τιμολόγιο”, “Invoice” can all indicate an invoice number.
- “Cobro”, “Pago”, “Transferencia”, “Trf”, “Bank” → Payment
- “Abono”, “Nota de crédito”, “Crédito”, “Πίστωση” → Credit Note
- “Factura”, “Τιμολόγιο”, “Invoice”, “Παρ.” → Invoice

If both DEBE and HABER missing but TOTAL exists → treat TOTAL as Debit (Invoice).
Ignore SALDO, IVA, Asiento, or summary lines.

Output strictly JSON array (no text), e.g.:
[
  {{
    "Alternative Document": "Invoice or reference number",
    "Date": "dd/mm/yy or yyyy-mm-dd",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "Invoice amount",
    "Credit": "Payment/Credit amount"
  }}
]

Text:
{text_block}
"""

        data = []
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                content = response.choices[0].message.content.strip()
                if i == 0:
                    st.text_area(f"🧠 GPT Response (Batch 1 – {model})", content, height=250, key=f"debug_{model}")
                data = parse_gpt_response(content, i // BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error ({model}): {e}")

        if not data:
            continue

        # --- Post-process records ---
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            if not alt_doc:
                continue

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            reason = str(row.get("Reason", "")).strip()

            # Cleanup and classification logic
            if debit_val and not credit_val:
                reason = "Invoice"
            elif credit_val and not debit_val:
                # Credit or payment
                if re.search(r"abono|nota|crédit|descuento|πίστωση|credit", str(row), re.IGNORECASE):
                    reason = "Credit Note"
                else:
                    reason = "Payment"
            elif debit_val == "" and credit_val == "":
                continue

            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": reason,
                "Debit": debit_val,
                "Credit": credit_val
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
    with st.spinner("📄 Extracting text + OCR fallback..."):
        lines, ocr_pages = extract_text_with_ocr(uploaded_pdf)

    if len(lines) == 0:
        st.error("❌ No text detected. Ensure Tesseract OCR and Poppler are installed (spa, ell, eng).")
    else:
        st.success(f"✅ Found {len(lines)} text lines.")
        if ocr_pages:
            st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")

        st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

        if st.button("🤖 Run Hybrid Extraction", type="primary"):
            with st.spinner("Analyzing text with GPT..."):
                data = extract_with_gpt(lines)

            if data:
                df = pd.DataFrame(data)
                st.success(f"✅ Extraction complete — {len(df)} valid records found.")
                st.dataframe(df, use_container_width=True, hide_index=True)

                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)

                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net", f"{net:,.2f}")

                st.download_button(
                    "⬇️ Download Excel",
                    data=to_excel_bytes(data),
                    file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ No structured data found in GPT output.")
else:
    st.info("Upload a vendor statement PDF to begin.")
