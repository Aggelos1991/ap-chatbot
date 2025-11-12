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
                # OCR fallback for pages without readable text
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
# GPT EXTRACTOR — Enhanced + Auto-Retry + Código ICN exclusion
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to detect Debit (DEBE) and Credit (HABER) or fallback TOTAL lines."""
    BATCH_SIZE = 60
    all_records = []
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        prompt = """
You are a financial data extractor for Spanish/Greek vendor statements.
Extract from lines:
- Date (Fecha)
- Doc num (Documento/N° DOC/Αρ. Παραστατικού/Αρ. Τιμολογίου or embedded in Concepto/Περιγραφή/Comentario as fallback)
- Description (Concepto/Περιγραφή/Comentario)
- Debit (DEBE/Χρέωση or TOTAL/ΤΕΛΙΚΟ/ΣΥΝΟΛΟ as fallback if no DEBE/HABER)
- Credit (HABER/Πίστωση)
Ignore SALDO, 'Asiento', 'Saldo', 'IVA', 'Total Saldo', "Código IC N".

Reason classification:
- Invoice: "Fra.", "Factura", "Τιμολόγιο", "Παραστατικό"
- Payment: "Cobro", "Pago", "Transferencia", "Remesa", "Bank", "Trf", "Pagado"
- Credit Note: "Abono", "Nota de crédito", "Crédito", "Descuento", "Πίστωση"

Output JSON array only:
[
  {
    "Alternative Document": "doc ref",
    "Date": "dd/mm/yy or yyyy-mm-dd",
    "Reason": "Invoice|Payment|Credit Note",
    "Debit": "amount",
    "Credit": "amount"
  }
]

Examples:
"31/01/25 1 245 N.F. A250213 NF A25021 907,98 6.355,74" → {"Alternative Document": "NF A25021", "Date": "31/01/25", "Reason": "Invoice", "Debit": "907,98", "Credit": ""}
"26/02/25 1 801 Cobro factura A250269 Rec NF A25069 542,90 3.719,83" → {"Alternative Document": "NF A25069", "Date": "26/02/25", "Reason": "Payment", "Debit": "", "Credit": "542,90"}
"Fecha: 15/03/25 Factura FRA-123 Total: 1.234,56" → {"Alternative Document": "FRA-123", "Date": "15/03/25", "Reason": "Invoice", "Debit": "1.234,56", "Credit": ""}
"10/04/25 Nota de crédito por devolución NC-456 0,00 789.12" → {"Alternative Document": "NC-456", "Date": "10/04/25", "Reason": "Credit Note", "Debit": "", "Credit": "789.12"}
"2025-05-20 Τιμολόγιο Αρ. Παραστατικού: ΤΙΜ-789 Χρέωση: 1.500,00" → {"Alternative Document": "ΤΙΜ-789", "Date": "2025-05-20", "Reason": "Invoice", "Debit": "1.500,00", "Credit": ""}
"01/06/25 Pago por transferencia bancaria ref. TRF-101 2.345,67" → {"Alternative Document": "TRF-101", "Date": "01/06/25", "Reason": "Payment", "Debit": "", "Credit": "2.345,67"}
"20/07/25 Concepto: Factura embedded F-2025-07 en pago parcial Total: 4.567,89" → {"Alternative Document": "F-2025-07", "Date": "20/07/25", "Reason": "Invoice", "Debit": "4.567,89", "Credit": ""}
Text:
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
        # === Post-process records ===
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            # exclude "Código IC N" and variants
            if re.search(r"codigo\s*ic\s*n", alt_doc, re.IGNORECASE):
                continue
            if not alt_doc or re.search(r"(asiento|saldo|iva|total\s+saldo)", alt_doc, re.IGNORECASE):
                continue
            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            reason = row.get("Reason", "").strip()
            # SALDO or dual values cleanup
            if debit_val and credit_val:
                if reason.lower() in ["payment", "credit note"]:
                    debit_val = ""
                elif reason.lower() == "invoice":
                    credit_val = ""
                else:
                    if abs(debit_val - credit_val) < 0.01 or min(debit_val, credit_val) / max(debit_val, credit_val) < 0.3:
                        if debit_val < credit_val:
                            debit_val = ""
                        else:
                            credit_val = ""
            # Classification fix
            if debit_val and not credit_val:
                reason = "Invoice"
            elif credit_val and not debit_val:
                if re.search(r"abono|nota|crédit|descuento|πίστωση", str(row), re.IGNORECASE):
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
            st.dataframe(df, use_container_width=True, hide_index=True)
            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net", f"{net:,.2f}")
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
