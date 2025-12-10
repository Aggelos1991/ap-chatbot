import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro")

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

# ==========================================================
# PDF + OCR EXTRACTION
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    """Extract ALL text lines (excluding Saldo lines), using OCR fallback only if needed."""
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    ocr_pages = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and len(text.strip()) > 10:
                for line in text.split("\n"):
                    clean_line = " ".join(line.split())
                    if not clean_line.strip():
                        continue
                    if re.search(r"\bsaldo\b", clean_line, re.IGNORECASE):
                        continue
                    all_lines.append(clean_line)
            else:
                try:
                    ocr_pages.append(i)
                    images = convert_from_bytes(pdf_bytes, dpi=250, first_page=i, last_page=i)
                    ocr_text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean_line = " ".join(line.split())
                        if not clean_line.strip():
                            continue
                        if re.search(r"\bsaldo\b", clean_line, re.IGNORECASE):
                            continue
                        all_lines.append(clean_line)
                except Exception as e:
                    st.warning(f"OCR skipped for page {i}: {e}")
    if ocr_pages:
        st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")
    return all_lines

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

# ==========================================================
# GPT EXTRACTOR — Referencia-based logic
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a financial data extractor for Spanish vendor statements (Libro mayor). 
Each line has columns such as:
Fecha | Asiento | Documento | Libro | Descripción | Referencia | F. valor | Debe | Haber | Saldo

Extract one JSON object per transaction with these fields:
- Alternative Document → The value under Referencia (leave empty if none)
- Concepto → Combine Asiento, Documento, and Descripción columns
- Date → Fecha
- Reason → based on Referencia + Debit/Credit logic:
   • If Referencia is empty → "Payment"
   • If Referencia exists and Debe > 0 → "Invoice"
   • If Referencia exists and Haber > 0 → "Credit Note"
- Debit → Debe value (only if visible)
- Credit → Haber value (only if visible)

⚠️ Rules:
- Never guess missing Referencia. If not visible, leave Alternative Document empty.
- Do NOT infer numeric values from nearby text — only copy explicit Debe/Haber numbers.
- Ignore “Saldo” columns completely.
- Return pure JSON array only, no explanations.

Example:
Line: "01/01/2023 VEN / 6887 183 /383005976 V 230101183005951 FP 010123 230101183005951 01/01/2023 6.171,48"
→ {{"Alternative Document": "230101183005951", "Concepto": "VEN / 6887 183 /383005976", "Date": "01/01/2023", "Reason": "Invoice", "Debit": "6.171,48", "Credit": ""}}

Line: "02/01/2023 GRL / 16811 GRL /0 V 221207183000015 IR VARIOS 2 02/01/2023  840,95"
→ {{"Alternative Document": "", "Concepto": "GRL / 16811 GRL /0", "Date": "02/01/2023", "Reason": "Payment", "Debit": "", "Credit": "840,95"}}

Text to analyze:
{text_block}
"""

        # GPT request
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
                st.warning(f"❌ GPT error with {model}: {e}")
                data = []

        if not data:
            continue

        # Normalize
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            reason = str(row.get("Reason", "")).strip()

            all_records.append({
                "Alternative Document": alt_doc,
                "Concepto": str(row.get("Concepto", "")).strip(),
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
    with st.spinner("📄 Extracting text from all pages (with OCR fallback)..."):
        lines = extract_raw_lines(uploaded_pdf)

    st.success(f"✅ Found {len(lines)} lines of text (Saldo lines removed).")
    st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Analyzing with GPT models..."):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success(f"✅ Extraction complete — {len(df)} valid records found!")
            st.dataframe(df[["Alternative Document", "Date", "Concepto", "Reason", "Debit", "Credit"]],
                         use_container_width=True, hide_index=True)
            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net", f"{net:,.2f}")
            except Exception as e:
                st.error(f"Totals error: {e}")

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No structured data detected. Check GPT response above.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
