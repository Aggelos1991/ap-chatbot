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
                ocr_pages.append(i)
                try:
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
        st.warning(f"⚠️ Batch {batch_num}: No JSON found.")
        return []
    try:
        return json.loads(json_match.group(0))
    except:
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
Extract structured records from Spanish/Greek vendor statements.

IMPORTANT ABSOLUTE RULES:
- Document number = ONLY the value from the field "Referencia".
- Do NOT extract invoice numbers from description, Concepto or anywhere else.
- If Referencia is empty → reason = Payment.
- If Referencia has DEBE → Invoice.
- If Referencia has HABER → Credit Note.

Extract:
- Fecha
- Referencia
- Asiento
- Concepto / Descripción
- DEBE
- HABER

Output JSON array:
[
  {{
    "Fecha": "",
    "Referencia": "",
    "Asiento": "",
    "Concepto": "",
    "Debit": "",
    "Credit": ""
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
                )
                content = response.choices[0].message.content.strip()
                data = parse_gpt_response(content, i // BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error ({model}): {e}")

        if not data:
            continue

        # =====================================================
        # FINAL CLASSIFICATION — EXACTLY YOUR RULES
        # =====================================================
        for row in data:
            referencia = str(row.get("Referencia", "")).strip()

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))

            # 1️⃣ Referencia empty → Payment
            if referencia == "":
                reason = "Payment"

            # 2️⃣ Referencia exists + DEBE > 0 → Invoice
            elif debit_val not in ("", 0) and float(debit_val) > 0:
                reason = "Invoice"

            # 3️⃣ Referencia exists + HABER > 0 → Credit Note
            elif credit_val not in ("", 0) and float(credit_val) > 0:
                reason = "Credit Note"

            # 4️⃣ Fallback
            else:
                reason = "Payment"

            all_records.append({
                "Document": referencia,
                "Date": str(row.get("Fecha", "")),
                "Asiento": str(row.get("Asiento", "")),
                "Concepto": str(row.get("Concepto", "")),
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
    with st.spinner("Extracting text (OCR fallback)…"):
        lines = extract_raw_lines(uploaded_pdf)

    st.success(f"Found {len(lines)} lines.")
    st.text_area("Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Extraction", type="primary"):

        with st.spinner("Running GPT extractor…"):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success(f"{len(df)} records extracted!")
            st.dataframe(df, use_container_width=True, hide_index=True)

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name="datafalcon_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No data extracted.")
else:
    st.info("Upload a PDF to start.")
