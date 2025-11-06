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



# === Load environment ===
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
# OCR EXTRACTION
# ==========================================================
def extract_text_with_ocr(uploaded_pdf):
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
# UTILITIES
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
        st.warning(f"⚠️ Batch {batch_num}: No JSON found.\n{content[:200]}")
        return []
    try:
        return json.loads(json_match.group(0))
    except json.JSONDecodeError as e:
        st.warning(f"⚠️ Batch {batch_num}: JSON decode error → {e}")
        return []

# ==========================================================
# GPT EXTRACTION
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        text_block = "\n".join(lines[i:i + BATCH_SIZE])
        prompt = f"""
You are a multilingual financial data extractor for vendor statements (Spanish / Greek / English).

Extract for each line:
- Alternative Document
- Date
- Reason or CONCEPTO (Invoice | Payment | Credit Note)
- Debit
- Credit
- Balance

Rules:
- "Saldo" or "Balance" always = Balance column (not Credit).
- "Pago", "Cobro", "Transferencia", "Remesa" → Payment.
- "Abono", "Nota de crédito", "Crédit" → Credit Note.
- Return ONLY JSON array.
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

        # === Restored behavior for proper Credit values ===
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            if not alt_doc:
                continue

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            balance_val = normalize_number(row.get("Balance", ""))
            reason = str(row.get("Reason", "")).strip().lower()

            # --- Classification and side assignment ---
            if re.search(r"pago|cobro|transferencia|remesa|trf|pagado|bank", str(row), re.IGNORECASE):
                reason = "Payment"
            elif re.search(r"abono|nota\s*de\s*crédito|crédit|descuento|πίστωση", str(row), re.IGNORECASE):
                reason = "Credit Note"
            else:
                reason = "Invoice"

            # Correct column placement
            if reason == "Payment":
                if not credit_val and debit_val:
                    credit_val, debit_val = debit_val, 0
                elif debit_val and credit_val == "":
                    credit_val = debit_val
                    debit_val = 0
            elif reason in ["Invoice", "Credit Note"]:
                if not debit_val and credit_val:
                    debit_val, credit_val = credit_val, 0

            # Skip blanks
            if debit_val == "" and credit_val == "":
                continue

            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": reason.title(),
                "Debit": debit_val,
                "Credit": credit_val,
                "Balance": balance_val
            })

    return all_records

# ==========================================================
# EXPORT TO EXCEL
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT INTERFACE
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text + running OCR fallback..."):
        lines, ocr_pages = extract_text_with_ocr(uploaded_pdf)

    if not lines:
        st.error("❌ No text detected. Check that Tesseract OCR is installed and language packs (spa, ell, eng) are available.")
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
                valid_balances = df["Balance"].apply(pd.to_numeric, errors="coerce").dropna()
                final_balance = valid_balances.iloc[-1] if not valid_balances.empty else total_debit - total_credit

                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("📊 Final Balance", f"{final_balance:,.2f}")

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
