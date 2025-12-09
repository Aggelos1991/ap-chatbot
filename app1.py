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

PRIMARY_MODEL = "gpt-4.1-mini"
BACKUP_MODEL = "gpt-4.1"

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
# PDF EXTRACTION (NO OCR)
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:

            text = page.extract_text()
            if not text:
                continue

            for line in text.split("\n"):
                clean_line = " ".join(line.split())

                if not clean_line.strip():
                    continue
                if re.search(r"\bsaldo\b", clean_line, re.IGNORECASE):
                    continue

                all_lines.append(clean_line)

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

IMPORTANT RULES:
- Document number = ONLY "Referencia".
- Do NOT extract invoice numbers from description.
- If Referencia empty → Payment.
- If Referencia has DEBE > 0 → Invoice.
- If Referencia has HABER > 0 → Credit Note.

Extract:
- Fecha
- Referencia
- Asiento
- Concepto
- DEBE
- HABER

Output JSON array.
Text:
{text_block}
"""

        data = []
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}]
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
        # FINAL CLASSIFICATION — YOUR RULES
        # =====================================================
        for row in data:
            referencia = str(row.get("Referencia", "")).strip()
            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))

            if referencia == "":
                reason = "Payment"
            elif debit_val not in ("", 0) and float(debit_val) > 0:
                reason = "Invoice"
            elif credit_val not in ("", 0) and float(credit_val) > 0:
                reason = "Credit Note"
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
    with st.spinner("Extracting text… (NO OCR)"):
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
