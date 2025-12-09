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
    """Extract ALL text lines from every page of the PDF (excluding Saldo lines), using OCR fallback."""
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
    """Try to extract JSON from GPT output safely."""
    json_match = re.search(r'\[.*\]', content, re.DOTALL)
    if not json_match:
        st.warning(f"⚠️ Batch {batch_num}: No JSON found. First 300 chars:\n{content[:300]}")
        return []
    try:
        data = json.loads(json_match.group(0))
        return data
    except json.JSONDecodeError as e:
        st.warning(f"⚠️ Batch {batch_num}: JSON decode error → {e}")
        return []

# ==========================================================
# GPT EXTRACTOR — Simplified Referencia-based Classification
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to extract data with Referencia-based classification."""
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a financial data extractor specialized in Spanish and Greek vendor statements.

Extract from each line:
- Fecha (Date)
- Referencia (Reference/Document number from "Referencia" column - THIS IS KEY)
- Concepto / Descripción (description)
- DEBE / Debit (Invoice amount)
- HABER / Credit (Payment or credit note amount)

⚠️ CRITICAL RULES:
1. Look specifically for the "Referencia" column/field for document numbers.
2. If a line has NO Referencia value, leave "Referencia" as empty string "".
3. COMPLETELY IGNORE 'Saldo' column - it is NOT Credit and NOT Debit!
4. SALDO is the running balance - NEVER extract it as Credit or Debit.
5. Only extract ACTUAL DEBE (Debit) and HABER (Credit) values.
6. If a line only has a Saldo value and no real DEBE/HABER, leave both Debit and Credit EMPTY.
7. Ignore lines with 'Asiento', 'IVA', or 'Total Saldo'.
8. DEBE values go in Debit field.
9. HABER values go in Credit field.
10. Output strictly JSON array only, no explanations.

⚠️ SALDO WARNING:
- Typical column order: DEBE | HABER | SALDO
- The LAST number on a line is usually SALDO (running balance) - DO NOT USE IT!
- Only use numbers that are clearly in DEBE or HABER columns.

OUTPUT FORMAT:
[
  {{
    "Referencia": "string (from Referencia column, empty if none)",
    "Concepto": "description text",
    "Date": "dd/mm/yy or yyyy-mm-dd",
    "Debit": "DEBE amount ONLY (not Saldo!), empty if none",
    "Credit": "HABER amount ONLY (not Saldo!), empty if none"
  }}
]

Examples:
Line: "15/03/25 REF-123 Factura servicios 1.234,56 5.000,00"
(1.234,56 is DEBE, 5.000,00 is SALDO - ignore Saldo!)
→ {{"Referencia": "REF-123", "Concepto": "Factura servicios", "Date": "15/03/25", "Debit": "1.234,56", "Credit": ""}}

Line: "20/03/25 Transferencia bancaria 500,00 4.500,00"
(500,00 is HABER, 4.500,00 is SALDO - ignore Saldo!)
→ {{"Referencia": "", "Concepto": "Transferencia bancaria", "Date": "20/03/25", "Debit": "", "Credit": "500,00"}}

Line: "25/03/25 Ajuste contable 4.500,00"
(Only SALDO shown, no DEBE/HABER - leave both empty!)
→ {{"Referencia": "", "Concepto": "Ajuste contable", "Date": "25/03/25", "Debit": "", "Credit": ""}}

Text to analyze:
{text_block}
"""

        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.0
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

        # === Post-process with SIMPLIFIED RULES ===
        for row in data:
            referencia = str(row.get("Referencia", "")).strip()
            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            concepto = str(row.get("Concepto", "")).strip()
            date_val = str(row.get("Date", "")).strip()

            # Skip if no financial values
            if debit_val == "" and credit_val == "":
                continue

            # Skip unwanted lines
            if re.search(r"(asiento|saldo|iva|total\s+saldo)", referencia, re.IGNORECASE):
                continue
            if re.search(r"codigo\s*ic\s*n", referencia, re.IGNORECASE):
                continue

            # === SIMPLIFIED CLASSIFICATION RULES ===
            # Rule 1: Referencia empty → Payment
            # Rule 2: Referencia has value + Debit → Invoice
            # Rule 3: Referencia has value + Credit → Credit Note

            if referencia == "":
                # No referencia = Payment
                reason = "Payment"
                # Payments should be in Credit column
                if debit_val and not credit_val:
                    credit_val = debit_val
                    debit_val = ""
            elif debit_val and not credit_val:
                # Has referencia + Debit = Invoice
                reason = "Invoice"
            elif credit_val and not debit_val:
                # Has referencia + Credit = Credit Note
                reason = "Credit Note"
            elif debit_val and credit_val:
                # Both values - use larger one to determine
                if float(debit_val) >= float(credit_val):
                    reason = "Invoice"
                    credit_val = ""
                else:
                    reason = "Credit Note"
                    debit_val = ""
            else:
                continue

            all_records.append({
                "Referencia": referencia,
                "Concepto": concepto,
                "Date": date_val,
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
            st.dataframe(df[["Referencia", "Date", "Concepto", "Reason", "Debit", "Credit"]], use_container_width=True, hide_index=True)

            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit (Invoices)", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit (Payments + CN)", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net Balance", f"{net:,.2f}")
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
