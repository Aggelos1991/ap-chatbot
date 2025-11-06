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
# GPT EXTRACTION  (HARDENED VERSION)
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        text_block = "\n".join(lines[i:i + BATCH_SIZE])
        prompt = f"""
You are a multilingual financial data extractor specialized in **vendor statements** (Spanish / Greek / English).

⚙️ TASK:
Read each line carefully and extract ONLY valid transaction lines (Factura, Abono, Pago, Transferencia, Nota de crédito).
IGNORE accounting entries like "Asiento", "Diario", "Regularización", or summary lines.

For every valid transaction, return a JSON array of objects with:
- "Alternative Document" → the true document/invoice number.
- "Date" → the date appearing on the same or nearby line.
- "Reason" → one of ["Invoice", "Payment", "Credit Note"].
- "Debit"
- "Credit"
- "Balance"

📘 RULES:
- Skip lines that contain only “Asiento”, “Diario”, “Apertura”, “Regularización”, or “Saldo anterior”.
- The document number can appear after words like "Num", "Núm.", "Documento", "Factura", "FAC", "FV", "CO", "AB", "Doc.", or inside "Concepto" or comments like "por factura 12345".
- Reject values like "Asiento 204", "Remesa", "Pago", "Transferencia" if they are **not** followed by a real invoice/credit reference.
- Prefer any code matching patterns: (F|FV|CO|AB|FA|FAC)\d{{3,}} or at least 5 consecutive digits.
- If text mentions "Pago", "Cobro", "Transferencia", "Remesa" → Reason = "Payment".
- If text mentions "Abono", "Nota de crédito", "Crédit", "Descuento", "Πίστωση" → Reason = "Credit Note".
- Otherwise, assume "Invoice".
- Do NOT invent any fields.
- Return only a pure JSON array, nothing else.

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

        # === Data post-processing ===
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            if not alt_doc:
                continue

            # Skip invalid refs
            if re.search(r"asiento|diario|apertura|regularizaci", alt_doc, re.IGNORECASE):
                continue
            if re.match(r"^(pago|remesa|transfer|trf|bank)$", alt_doc, re.IGNORECASE):
                continue

            # Force-detect real doc number
            match = re.search(r"((F|FV|CO|AB|FAC|FA)\d{3,}|\d{5,})", alt_doc)
            if match:
                alt_doc = match.group(1)
            else:
                continue

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            balance_val = normalize_number(row.get("Balance", ""))
            reason = str(row.get("Reason", "")).strip().lower()

            # Correct reason detection
            if re.search(r"pago|cobro|transfer|remesa|trf|bank", str(row), re.IGNORECASE):
                reason = "Payment"
            elif re.search(r"abono|nota\s*de\s*crédito|crédit|descuento|πίστωση", str(row), re.IGNORECASE):
                reason = "Credit Note"
            else:
                reason = "Invoice"

            # Fix debit/credit placement
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
