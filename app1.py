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
# HELPERS
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

def extract_raw_lines(uploaded_pdf):
    """Extracts all lines from PDF, automatically running OCR if no text layer is found."""
    all_lines, ocr_pages = [], []
    pdf_bytes = uploaded_pdf.read()

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and len(text.strip()) > 20:
                for line in text.split("\n"):
                    clean_line = " ".join(line.split())
                    if clean_line:
                        all_lines.append(clean_line)
            else:
                # --- OCR fallback ---
                ocr_pages.append(i)
                image = convert_from_bytes(pdf_bytes, dpi=200, first_page=i, last_page=i)[0]
                ocr_text = pytesseract.image_to_string(image, lang="spa+ell+eng")
                for line in ocr_text.split("\n"):
                    clean_line = " ".join(line.split())
                    if clean_line:
                        all_lines.append(clean_line)

    return all_lines, ocr_pages

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
# GPT EXTRACTOR
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a financial data extractor specialized in Spanish and Greek vendor statements.

Each line may contain:
- Fecha (Date)
- Documento / N° DOC / Αρ. Παραστατικού / Αρ. Τιμολογίου (Document number)
- Concepto / Περιγραφή / Comentario (description)
- DEBE / Χρέωση (Invoice amount)
- HABER / Πίστωση (Payments or credit notes)
- TOTAL lines (if no DEBE/HABER)
Follow the same extraction logic as before.

OUTPUT JSON ONLY:
[
  {{
    "Alternative Document": "string",
    "Date": "dd/mm/yy or yyyy-mm-dd",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "DEBE or TOTAL amount",
    "Credit": "HABER amount"
  }}
]
Text:
{text_block}
"""
        data = []
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.0
                )
                content = resp.choices[0].message.content.strip()
                if i == 0:
                    st.text_area(f"🧠 GPT Response (Batch 1 – {model})", content, height=250, key=f"debug_{model}")
                data = parse_gpt_response(content, i // BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error ({model}): {e}")
                data = []
        if not data:
            continue

        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            if re.search(r"codigo\s*ic\s*n", alt_doc, re.I): 
                continue
            if not alt_doc or re.search(r"(asiento|saldo|iva|total\s+saldo)", alt_doc, re.I):
                continue

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            reason = row.get("Reason", "").strip()

            if debit_val and not credit_val:
                reason = "Invoice"
            elif credit_val and not debit_val:
                if re.search(r"abono|nota|crédit|descuento|πίστωση", str(row), re.I):
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
    with st.spinner("📄 Extracting text + OCR from pages..."):
        lines, ocr_pages = extract_raw_lines(uploaded_pdf)

    st.success(f"✅ {len(lines)} lines extracted.")
    if ocr_pages:
        st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")

    st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Hybrid GPT Extraction", type="primary"):
        with st.spinner("Analyzing text with GPT..."):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success(f"✅ Extraction complete — {len(df)} records found!")
            st.dataframe(df, use_container_width=True, hide_index=True)

            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)
                c1, c2, c3 = st.columns(3)
                c1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                c2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                c3.metric("⚖️ Net", f"{net:,.2f}")
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
