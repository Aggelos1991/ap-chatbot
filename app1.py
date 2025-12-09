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
        prompt = f"""
You are a financial data extractor specialized in Spanish and Greek vendor statements.
Each line may contain:
- Fecha (Date)
- Documento / N° DOC / Αρ. Παραστατικού / Αρ. Τιμολογίου (Document number)
- Concepto / Περιγραφή / Comentario (description)
- DEBE / Χρέωση (Invoice amount)
- HABER / Πίστωση (Payments or credit notes)
- SALDO (ignore)
- TOTAL / TOTALES / ΤΕΛΙΚΟ / ΣΥΝΟΛΟ / IMPORTE TOTAL / TOTAL FACTURA — treat as invoice total if no DEBE/HABER available
⚠️ RULES
1. Ignore lines with 'Asiento', 'Saldo', 'IVA', or 'Total Saldo'.
2. Exclude codes like "Código IC N" or similar from document detection.
3. If "N° DOC" or "Documento" missing, detect invoice-like code (FAC123, F23, INV-2024, FRA-005, ΤΙΜ 123, etc or embedded in Concepto/Περιγραφή/Comentario as fallback)).
4. Detect reason:
   - "Cobro", "Pago", "Transferencia", "Remesa", "Bank", "Trf", "Pagado" → Payment
   - "Abono", "Nota de crédito", "Crédito", "Descuento", "Πίστωση" → Credit Note
   - "Fra.", "Factura", "Τιμολόγιο", "Παραστατικό" → Invoice
5. DEBE / Χρέωση → Invoice (put in Debit)
6. HABER / Πίστωση → Payment or Credit Note (put in Credit)
7. If neither DEBE nor HABER exists but TOTAL/TOTALES/ΤΕΛΙΚΟ/ΣΥΝΟΛΟ appear, use that value as Debit (Invoice total).
8. Output strictly JSON array only, no explanations.
Examples:
Line: "31/01/25 1 245 N.F. A250213 NF A25021 907,98 6.355,74"
Output object: {{"Alternative Document": "NF A25021", "Date": "31/01/25", "Reason": "Invoice", "Debit": "907,98", "Credit": ""}}
Line: "26/02/25 1 801 Cobro factura A250269 Rec NF A25069 542,90 3.719,83"
Output object: {{"Alternative Document": "NF A25069", "Date": "26/02/25", "Reason": "Payment", "Debit": "", "Credit": "542,90"}}
Line: "Fecha: 15/03/25 Factura FRA-123 Total: 1.234,56"
Output object: {{"Alternative Document": "FRA-123", "Date": "15/03/25", "Reason": "Invoice", "Debit": "1.234,56", "Credit": ""}}
OUTPUT FORMAT:
[
  {{
    "Alternative Document": "string (invoice or payment ref)",
    "Concepto": "factura num from description",
    "Date": "dd/mm/yy or yyyy-mm-dd",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "DEBE or TOTAL amount",
    "Credit": "HABER amount"
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
            # Classification fix + negative DEBE handling
           # Classification fix + negative DEBE handling
            # === Classification fix + safe negative DEBE handling ===
            if debit_val != "" and credit_val == "":
                try:
                    val = float(debit_val)
                    if val < 0:
                        # Negative DEBE → move to Credit
                        credit_val = round(abs(val), 2)
                        debit_val = ""
                        reason = "Credit Note"
                    else:
                        reason = "Invoice"
                except:
                    reason = "Invoice"
            
            elif credit_val != "" and debit_val == "":
                if re.search(r"abono|nota|crédit|descuento|πίστωση", str(row), re.IGNORECASE):
                    reason = "Credit Note"
                else:
                    reason = "Payment"
            
            elif debit_val == "" and credit_val == "":
                continue
            all_records.append({
                "Alternative Document": alt_doc,
                "Concepto": str(row.get("Concepto", "")).strip(),
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
            st.dataframe(df[["Alternative Document", "Date", "Concepto", "Reason", "Debit", "Credit"]], use_container_width=True, hide_index=True)
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
