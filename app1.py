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
# GPT EXTRACTION (bulletproofed)
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        text_block = "\n".join(lines[i:i + BATCH_SIZE])
        prompt = f"""
You are a multilingual financial data extractor for vendor statements (Spanish / Greek / English).

Extract for each line ONLY real transaction rows — invoices, payments, or credit notes.
IGNORE accounting rows such as "Asiento", "Diario", "Regularización", "Saldo anterior", "Cuenta 43...".
If you are unsure, skip the line.

For each valid line, return:
- "Alternative Document": real document/invoice number (from Num., Documento, Factura, or Concepto). 
  It often looks like A741387, AB0718, FV12345, FAC2345, CO1234, or any code with ≥5 digits.
  NEVER use account codes like 43xxxxxx or text like "Asiento".
- "Date": the date on that line (dd/mm/yyyy or yyyy-mm-dd)
- "Reason": one of ["Invoice", "Payment", "Credit Note"]
- "Debit"
- "Credit"
- "Balance"

Rules:
- Skip lines containing "Asiento", "Diario", "Apertura", "Regularización", or "Saldo anterior".
- "Pago", "Cobro", "Transferencia", "Remesa", "Orden de cobro", "BNKI", "Banco" → Payment
- "Abono", "Reversión", "Nota de crédito", "Crédit", "Πίστωση" → Credit Note
- Otherwise → Invoice
- "Saldo" or "Balance" always means the balance column (not Credit).
- Document numbers appear after words like Num., Número, Documento, Factura, Fact., Doc., or in "Concepto" / "Comentarios".
- Return only valid JSON array, no explanations or text.

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
            if not alt_doc or re.match(r"^(43\d{6,}|asiento|diario|regularizaci)", alt_doc, re.IGNORECASE):
                continue

            # force revalidate doc number
            match = re.search(r"((A|AB|AC|FV|FAC|FA|CO)\d{3,}|\d{5,})", alt_doc)
            if match:
                alt_doc = match.group(1)
            else:
                continue

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            balance_val = normalize_number(row.get("Balance", ""))
            reason = str(row.get("Reason", "")).strip().lower()

            if re.search(r"pago|cobro|transferencia|remesa|bnki|banco|trf|pagado|bank|orden\s+de\s+cobro", str(row), re.IGNORECASE):
                reason = "Payment"
            elif re.search(r"abono|nota\s*de\s*cr[eé]dito|cr[eé]dit|reversi[oó]n|πίστωση", str(row), re.IGNORECASE):
                reason = "Credit Note"
            else:
                reason = "Invoice"

            if reason == "Payment":
                if debit_val and not credit_val:
                    credit_val, debit_val = debit_val, 0
            elif reason in ["Invoice", "Credit Note"]:
                if credit_val and not debit_val:
                    debit_val, credit_val = credit_val, 0

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
