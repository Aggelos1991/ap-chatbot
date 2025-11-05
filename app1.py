import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT+OCR Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Hybrid GPT + OCR Extractor")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

# ===== AUTH FIX (PROJECT-BASED API) =====
api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
organization = os.getenv("OPENAI_ORG_ID") or st.secrets.get("OPENAI_ORG_ID")
project = os.getenv("OPENAI_PROJECT_ID") or st.secrets.get("OPENAI_PROJECT_ID")

if not api_key:
    st.error("❌ No OpenAI API key found. Add it to .env or Streamlit Secrets.")
    st.stop()

client = OpenAI(
    api_key=api_key,
    organization=organization,
    project=project
)

PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"

# ==========================================================
# OCR-ENHANCED TEXT EXTRACTION
# ==========================================================
def extract_text_with_ocr(uploaded_pdf):
    """Extract text from PDF using both pdfplumber and OCR fallback."""
    all_lines, ocr_pages = [], []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and len(text.strip()) > 10:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean:
                        all_lines.append(clean)
            else:
                # OCR fallback
                ocr_pages.append(i)
                try:
                    img = convert_from_bytes(pdf_bytes, dpi=250, first_page=i, last_page=i)[0]
                    ocr_text = pytesseract.image_to_string(img, lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean:
                            all_lines.append(clean)
                except Exception as e:
                    st.warning(f"OCR skipped for page {i}: {e}")

    return all_lines, ocr_pages

# ==========================================================
# GPT EXTRACTION (SALDO FIXED)
# ==========================================================
def normalize_number(value):
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
    """Multilingual GPT extraction tuned for Spanish/Greek/English statements with Balance column."""
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        prompt = f"""
You are a multilingual financial data extractor for vendor statements (Spanish / Greek / English).

For each line, extract:
- Document / Reference / Invoice number (Documento, N° DOC, Αρ. Παραστατικού, Reference, Fra., Τιμολόγιο)
- Date (Fecha, Ημερομηνία)
- Reason (Invoice | Payment | Credit Note)
- Debit (DEBE / Χρέωση / TOTAL / ΣΥΝΟΛΟ when DEBE/HABER missing)
- Credit (HABER / Πίστωση)
- Balance (Saldo / Υπόλοιπο / Συνολικό Υπόλοιπο / Balance)

Rules:
- "SALDO", "Υπόλοιπο", or "Balance" always represent the BALANCE column — never Credit.
- Ignore headers like Asiento, IVA, Total Saldo, or empty lines.
- If line text contains 'Cobro', 'Pago', 'Transferencia', 'Remesa', classify it as Payment (even if Reason missing).
- If text contains 'Abono', 'Nota de crédito', 'Crédit', or 'Πίστωση', classify as Credit Note.
- TOTAL/TOTALES/ΣΥΝΟΛΟ used only if DEBE/HABER missing.

Return valid JSON only:
[
  {{
    "Alternative Document": "Invoice or reference number",
    "Date": "dd/mm/yy or yyyy-mm-dd",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "number",
    "Credit": "number",
    "Balance": "number"
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

        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            if not alt_doc:
                continue

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            balance_val = normalize_number(row.get("Balance", ""))
            reason = str(row.get("Reason", "")).strip()

            if re.search(r"cobro|pago|transferencia|remesa", str(row), re.IGNORECASE):
                reason = "Payment"
            elif re.search(r"abono|nota de crédito|crédit|πίστωση", str(row), re.IGNORECASE):
                reason = "Credit Note"
            elif not reason:
                reason = "Invoice"

            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": reason,
                "Debit": debit_val,
                "Credit": credit_val,
                "Balance": balance_val
            })

    return all_records

# ==========================================================
# EXPORT UTIL
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
    with st.spinner("📄 Extracting text + running OCR fallback..."):
        lines, ocr_pages = extract_text_with_ocr(uploaded_pdf)

    if len(lines) == 0:
        st.error("❌ No text detected. Make sure Tesseract OCR is installed and language packs (spa, ell, eng) are available.")
    else:
        st.success(f"✅ Found {len(lines)} lines of text!")
        if ocr_pages:
            st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")

        st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

        if st.button("🤖 Run Hybrid Extraction", type="primary"):
            with st.spinner("Analyzing with GPT..."):
                data = extract_with_gpt(lines)

            if data:
                df = pd.DataFrame(data)
                st.success(f"✅ Extraction complete — {len(df)} valid records found!")
                st.dataframe(df, use_container_width=True, hide_index=True)

                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                total_balance = df["Balance"].apply(pd.to_numeric, errors="coerce").dropna().iloc[-1] if df["Balance"].notna().any() else 0
                net = round(total_debit - total_credit, 2)

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net", f"{net:,.2f}")
                col4.metric("📊 Final Balance", f"{total_balance:,.2f}")

                st.download_button(
                    "⬇️ Download Excel",
                    data=to_excel_bytes(data),
                    file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ No structured data found in GPT output.")
else:
    st.info("Upload a PDF to begin.")
