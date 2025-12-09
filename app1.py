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
You are a financial data extractor for Spanish "Libro Mayor" (General Ledger) statements.

THE EXACT COLUMN ORDER IN THIS DOCUMENT IS:
Fecha | Asiento | Documento | Libro | Descripción | Referencia | F. valor | Debe | Haber | Saldo

⚠️ CRITICAL - COLUMN POSITIONS:
- Column 8 = DEBE (Debit) - extract to "Debit" field
- Column 9 = HABER (Credit) - extract to "Credit" field  
- Column 10 = SALDO (Balance) - IGNORE THIS COMPLETELY!

The numbers appear in this order: DEBE | HABER | SALDO
- If a row has ONE number before Saldo → determine if it's in Debe or Haber position
- If a row has TWO numbers before Saldo → first is Debe, second is Haber
- The LAST number on each line is almost always SALDO - DO NOT USE IT!

EXTRACT THESE FIELDS:
- Referencia: The reference number (column 6) - may be empty!
- Descripción: The description text
- F. valor: The value date (column 7)
- Debe: Amount from DEBE column ONLY
- Haber: Amount from HABER column ONLY

⚠️ RULES:
1. If Referencia column is EMPTY, leave "Referencia" as ""
2. DEBE values → put in "Debit" field
3. HABER values → put in "Credit" field
4. NEVER use SALDO values - they are just running totals!
5. If only one amount exists and no clear column indicator, check the SALDO: if the Saldo INCREASES, the value is DEBE; if Saldo DECREASES, the value is HABER
6. Skip header rows and total rows

OUTPUT FORMAT (strict JSON array):
[
  {{
    "Referencia": "reference number or empty string",
    "Descripcion": "description text",
    "Date": "dd/mm/yyyy",
    "Debit": "DEBE amount or empty",
    "Credit": "HABER amount or empty"
  }}
]

EXAMPLES FROM THIS DOCUMENT:
Line: "01/01/2023 VEN / 6887 183 /383005976 V 230101183005951 FP 010123 F 230101183005951 01/01/2023 6.171,48 7.488,96"
→ Referencia=230101183005951, Debe=6.171,48, Saldo=7.488,96 (ignore saldo)
→ {{"Referencia": "230101183005951", "Descripcion": "230101183005951 FP 010123 F", "Date": "01/01/2023", "Debit": "6.171,48", "Credit": ""}}

Line: "02/01/2023 GRL / 16811 GRL /0 V 221207183000015 IR VARIOS 2 02/01/2023 840,95 6.648,01"
→ Referencia=EMPTY (no ref number), Haber=840,95, Saldo=6.648,01 (ignore saldo)
→ {{"Referencia": "", "Descripcion": "221207183000015 IR VARIOS 2", "Date": "02/01/2023", "Debit": "", "Credit": "840,95"}}

Line: "26/01/2023 VEN / 22339 938 /338000858 V 2294126 230126938000024 FP 230126938000024 26/01/2023 580,00 -580,00"
→ Referencia=230126938000024, Haber=580,00 (because Saldo went negative/decreased)
→ {{"Referencia": "230126938000024", "Descripcion": "2294126 230126938000024 FP", "Date": "26/01/2023", "Debit": "", "Credit": "580,00"}}

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
            descripcion = str(row.get("Descripcion", "") or row.get("Concepto", "")).strip()
            date_val = str(row.get("Date", "")).strip()

            # Skip if no financial values
            if debit_val == "" and credit_val == "":
                continue

            # Skip unwanted lines
            if re.search(r"(total\s+saldo|total\s+debe|total\s+haber)", descripcion, re.IGNORECASE):
                continue

            # === CLASSIFICATION RULES ===
            # Rule 1: No Referencia → Payment (value in Credit/Haber)
            # Rule 2: Referencia + Debit (DEBE) → Invoice
            # Rule 3: Referencia + Credit (HABER) → Credit Note

            if referencia == "":
                # No referencia = Payment
                reason = "Payment"
            elif debit_val and not credit_val:
                # Has referencia + Debit (DEBE) = Invoice
                reason = "Invoice"
            elif credit_val and not debit_val:
                # Has referencia + Credit (HABER) = Credit Note
                reason = "Credit Note"
            elif debit_val and credit_val:
                # Both values present - classify by which is larger
                if float(debit_val) >= float(credit_val):
                    reason = "Invoice"
                else:
                    reason = "Credit Note"
            else:
                continue

            all_records.append({
                "Referencia": referencia,
                "Descripcion": descripcion,
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
            st.dataframe(df[["Referencia", "Date", "Descripcion", "Reason", "Debit", "Credit"]], use_container_width=True, hide_index=True)

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
