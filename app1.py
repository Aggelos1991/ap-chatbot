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
# PDF TABLE EXTRACTION - Direct column parsing
# ==========================================================
def extract_table_data(uploaded_pdf):
    """Extract table data directly using pdfplumber's table extraction."""
    all_records = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # Try to extract tables
            tables = page.extract_tables()
            
            if tables:
                for table in tables:
                    for row in table:
                        if not row or len(row) < 5:
                            continue
                        # Skip header rows
                        row_text = " ".join(str(cell or "") for cell in row).lower()
                        if "fecha" in row_text or "debe" in row_text or "haber" in row_text:
                            continue
                        if "total" in row_text or "saldo anterior" in row_text:
                            continue
                        
                        all_records.append(row)
            else:
                # Fallback: extract text and parse manually
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        clean_line = " ".join(line.split())
                        if clean_line:
                            all_records.append(clean_line)
    
    return all_records


def parse_libro_mayor_row(row):
    """
    Parse a row from Libro Mayor table.
    Expected columns: Fecha | Asiento | Documento | Libro | Descripción | Referencia | F. valor | Debe | Haber | Saldo
    """
    if isinstance(row, str):
        return None  # Skip text lines, handle separately
    
    if len(row) < 10:
        return None
    
    # Column indices (0-based)
    # 0=Fecha, 1=Asiento, 2=Documento, 3=Libro, 4=Descripción, 5=Referencia, 6=F.valor, 7=Debe, 8=Haber, 9=Saldo
    
    fecha = str(row[0] or "").strip()
    descripcion = str(row[4] or "").strip()
    referencia = str(row[5] or "").strip()
    f_valor = str(row[6] or "").strip()
    debe_raw = str(row[7] or "").strip()
    haber_raw = str(row[8] or "").strip()
    
    # Skip if no date (header or invalid row)
    if not fecha or not re.match(r"\d{2}/\d{2}/\d{4}", fecha):
        return None
    
    # Normalize amounts
    debe = normalize_number(debe_raw) if debe_raw else ""
    haber = normalize_number(haber_raw) if haber_raw else ""
    
    # Skip if no financial values
    if debe == "" and haber == "":
        return None
    
    # Classification based on your rules:
    # - No Referencia → Payment
    # - Referencia + Debe → Invoice  
    # - Referencia + Haber → Credit Note
    
    if referencia == "" or referencia == "0" or referencia.lower() == "null":
        referencia = ""
        reason = "Payment"
    elif debe and not haber:
        reason = "Invoice"
    elif haber and not debe:
        reason = "Credit Note"
    elif debe and haber:
        if float(debe) >= float(haber):
            reason = "Invoice"
        else:
            reason = "Credit Note"
    else:
        return None
    
    return {
        "Referencia": referencia,
        "Descripcion": descripcion,
        "Date": f_valor if f_valor else fecha,
        "Reason": reason,
        "Debit": debe,
        "Credit": haber
    }


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
# GPT EXTRACTOR — Fallback for non-table PDFs
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to extract data when table structure is not detected."""
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a financial data extractor for Spanish vendor statements.

Extract from each line:
- Referencia: Document/reference number (if present, otherwise empty "")
- Descripcion: Description text
- Date: Transaction date
- Debit: DEBE amount (invoices)
- Credit: HABER amount (payments/credits)

IGNORE the SALDO column (running balance) - it's usually the LAST number on each line.

OUTPUT FORMAT (strict JSON array):
[
  {{
    "Referencia": "reference number or empty string",
    "Descripcion": "description text",
    "Date": "dd/mm/yyyy",
    "Debit": "amount or empty",
    "Credit": "amount or empty"
  }}
]

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

        for row in data:
            referencia = str(row.get("Referencia", "")).strip()
            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            descripcion = str(row.get("Descripcion", "")).strip()
            date_val = str(row.get("Date", "")).strip()

            if debit_val == "" and credit_val == "":
                continue

            # Classification
            if referencia == "":
                reason = "Payment"
            elif debit_val and not credit_val:
                reason = "Invoice"
            elif credit_val and not debit_val:
                reason = "Credit Note"
            else:
                reason = "Invoice" if float(debit_val or 0) >= float(credit_val or 0) else "Credit Note"

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
    with st.spinner("📄 Extracting table data from PDF..."):
        raw_data = extract_table_data(uploaded_pdf)
    
    st.success(f"✅ Found {len(raw_data)} rows in PDF.")
    
    # Check if we got table data or text lines
    has_tables = any(isinstance(row, list) for row in raw_data)
    
    if has_tables:
        st.info("📊 Table structure detected - using direct column extraction")
        
        if st.button("🔍 Extract Data", type="primary"):
            records = []
            for row in raw_data:
                parsed = parse_libro_mayor_row(row)
                if parsed:
                    records.append(parsed)
            
            if records:
                df = pd.DataFrame(records)
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
                    data=to_excel_bytes(records),
                    file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ No valid records found in table.")
    else:
        st.info("📝 No table structure found - will use GPT extraction")
        lines = [row for row in raw_data if isinstance(row, str)]
        st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)
        
        if st.button("🤖 Run GPT Extraction", type="primary"):
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
