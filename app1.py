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
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    ocr_pages = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()

            if text and len(text.strip()) > 10:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if not clean:
                        continue
                    if re.search(r"\bsaldo\b", clean, re.IGNORECASE):
                        continue
                    all_lines.append(clean)
            else:
                ocr_pages.append(i)
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=250, first_page=i, last_page=i)
                    ocr_text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")

                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if not clean:
                            continue
                        if re.search(r"\bsaldo\b", clean, re.IGNORECASE):
                            continue
                        all_lines.append(clean)

                except Exception as e:
                    st.warning(f"OCR skipped on page {i}: {e}")

    if ocr_pages:
        st.info(f"OCR applied on pages: {ocr_pages}")

    return all_lines

# ==========================================================
# JSON PARSER
# ==========================================================
def parse_gpt_response(content, batch_num):
    match = re.search(r'\[.*\]', content, re.DOTALL)
    if not match:
        st.warning(f"No JSON found in batch {batch_num}")
        return []
    try:
        return json.loads(match.group(0))
    except:
        st.warning(f"JSON decode error in batch {batch_num}")
        return []

# ==========================================================
# GPT EXTRACTOR — FINAL VERSION (Referencia ONLY)
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 60
    final_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
Extract all accounting entries. ONLY return these fields:

- Referencia (from explicit keywords only: Ref, Referencia, N°, Doc, Documento, Num.)
- Concepto
- Date
- Debit (DEBE)
- Credit (HABER)

RULES:
1. Do NOT invent invoice numbers.
2. If no Referencia is present → leave it blank.
3. If Debit has value → Reason = "Invoice".
4. If Credit has value → Reason = "Credit Note" unless description indicates payment.
5. If Credit and description includes: pago, cobro, transferencia, bank, trf → Reason = "Payment".
6. JSON array ONLY.

Text:
{text_block}
"""

        # Try models
        data = []
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                r = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                content = r.choices[0].message.content.strip()

                if i == 0:
                    st.text_area("GPT DEBUG Response", content, height=250)

                data = parse_gpt_response(content, i)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error: {e}")
                data = []

        if not data:
            continue

        # ============================
        # FINAL POST PROCESSING
        # ============================
        for row in data:

            referencia = str(row.get("Referencia", "")).strip()
            concepto = str(row.get("Concepto", "")).strip()
            date = str(row.get("Date", "")).strip()

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))

            # === CLASSIFICATION ===
            if debit_val and not credit_val:
                reason = "Invoice"
            elif credit_val and not debit_val:
                if re.search(r"pago|cobro|transfer|bank|trf", concepto, re.IGNORECASE):
                    reason = "Payment"
                else:
                    reason = "Credit Note"
            else:
                continue  # discard useless rows

            final_records.append({
                "Referencia": referencia,
                "Concepto": concepto,
                "Date": date,
                "Reason": reason,
                "Debit": debit_val,
                "Credit": credit_val
            })

    return final_records

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
    with st.spinner("Extracting PDF + OCR…"):
        lines = extract_raw_lines(uploaded_pdf)

    st.success(f"Extracted {len(lines)} lines.")

    st.text_area("Preview (first 30):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Extraction"):
        with st.spinner("Processing with GPT…"):
            records = extract_with_gpt(lines)

        if records:
            df = pd.DataFrame(records)
            st.success(f"Extracted {len(df)} valid rows!")

            st.dataframe(df, use_container_width=True, hide_index=True)

            # Totals
            total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
            total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Debit", f"{total_debit:,.2f}")
            col2.metric("Total Credit", f"{total_credit:,.2f}")
            col3.metric("Net", f"{total_debit - total_credit:,.2f}")

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(records),
                file_name="statement_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No valid structured data extracted.")

else:
    st.info("Upload a PDF to begin.")
