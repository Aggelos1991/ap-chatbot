import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ====================== NEW OCR IMPORTS ======================
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
# ==========================================================

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("DataFalcon Pro")
try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass
api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("No OpenAI API key found. Add it to .env or Streamlit Secrets.")
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

# ====================== HYBRID TEXT EXTRACTOR (OCR + TEXT) ======================
def extract_raw_lines(uploaded_pdf):
    """Extract ALL text lines from every page – OCR for scanned PDFs."""
    all_lines = []
    pdf_bytes = uploaded_pdf.getvalue()

    # 1. Try fast text extraction
    with pdfplumber.open(uploaded_pdf) as pdf:
        sample_text = any(page.extract_text() for page in pdf.pages[:3] if page.extract_text())

    if sample_text:
        st.info("Detected searchable PDF → using fast text extraction")
        with pdfplumber.open(uploaded_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split("\n"):
                    clean_line = " ".join(line.split())
                    if clean_line.strip():
                        all_lines.append(clean_line)
    else:
        # 2. OCR MODE – Scanned PDF
        st.warning("No text layer found → switching to OCR (slower, 1–3 min)")
        with st.spinner("Running OCR on every page..."):
            try:
                images = convert_from_bytes(pdf_bytes, dpi=300, fmt="png", thread_count=4)
                for i, img in enumerate(images):
                    with st.status(f"OCR Page {i+1}/{len(images)}") as status:
                        status.write(f"Reading page {i+1}…")
                        text = pytesseract.image_to_string(
                            img,
                            lang='spa',  # Spanish
                            config='--psm 6'  # Assume a uniform block of text
                        )
                        for line in text.split("\n"):
                            clean_line = " ".join(line.split())
                            if clean_line.strip():
                                all_lines.append(clean_line)
                        status.update(label=f"Page {i+1} completed", state="complete")
                st.success(f"OCR finished → {len(all_lines)} lines extracted!")
            except Exception as e:
                st.error(f"OCR failed: {e}")
                st.info("Install Tesseract OCR and add it to PATH.")
                return []

    return all_lines
# ===========================================================================

def parse_gpt_response(content, batch_num):
    """Try to extract JSON from GPT output safely."""
    json_match = re.search(r'\[.*\]', content, re.DOTALL)
    if not json_match:
        st.warning(f"Batch {batch_num}: No JSON found. First 300 chars:\n{content[:300]}")
        return []
    try:
        data = json.loads(json_match.group(0))
        return data
    except json.JSONDecodeError as e:
        st.warning(f"Batch {batch_num}: JSON decode error → {e}")
        return []

# ==========================================================
# GPT EXTRACTOR — Enhanced + Auto-Retry
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to detect Debit (DEBE) and Credit (HABER) from vendor statements."""
    BATCH_SIZE = 60
    all_records = []
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        prompt = f"""
You are a financial data extractor specialized in Spanish vendor statements.
Each line contains:
- Fecha (Date)
- N° DOC or Documento (Document number)
- Comentario / Concepto / Descripción (may contain invoice or payment details)
- DEBE (Invoice amounts)
- HABER / CRÉDITO (Payments or credit notes)
- SALDO (running balance — IGNORE)
Your task: extract all valid transactions and classify them precisely.
CLASSIFICATION RULES:
1. Ignore lines with Asiento, Saldo, IVA, or Total.
2. If "N° DOC" missing, find invoice-like pattern (FRA 209, FAC1234, FACTURA 1775, INV-2024-01).
3. Use Comentario to detect the reason:
   - Cobro, Pago, Transferencia, Remesa, Bank, Trf, Pagado → Payment
   - Abono, Nota de crédito, Crédito, Descuento → Credit Note
   - Fra., Factura, FRA, Factura Proveedor → Invoice
4. DEBE → Invoice
5. HABER → Payment or Credit Note
6. If both DEBE & HABER appear, keep only the correct side.
7. Never use SALDO.
8. Output strictly JSON array only, no explanations.
OUTPUT FORMAT:
[
  {{
    "Alternative Document": "...",
    "Date": "dd/mm/yy",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "DEBE amount or empty string",
    "Credit": "HABER amount or empty string"
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
                    st.text_area(f"GPT Response (Batch 1 – {model})", content, height=250, key=f"debug_{model}")
                data = parse_gpt_response(content, i // BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error with {model}: {e}")
                data = []
        if not data:
            continue

        # === Post-process records ===
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            if not alt_doc or re.search(r"(asiento|saldo|total|iva)", alt_doc, re.IGNORECASE):
                continue
            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            reason = row.get("Reason", "").strip()

            # SALDO / DOUBLE-SIDE CLEANUP
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

            # Classification correction
            if debit_val and not credit_val:
                reason = "Invoice"
            elif credit_val and not debit_val:
                if re.search(r"abono|nota|crédit|descuento", str(row), re.IGNORECASE):
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
uploaded_pdf = st.file_uploader("Upload Vendor Statement (PDF)", type=["pdf"])
if uploaded_pdf:
    with st.spinner("Extracting text from all pages..."):
        lines = extract_raw_lines(uploaded_pdf)
    st.success(f"Found {len(lines)} lines of text!")
    st.text_area("Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("Run Hybrid Extraction", type="primary"):
        with st.spinner("Analyzing with GPT models..."):
            data = extract_with_gpt(lines)
        if data:
            df = pd.DataFrame(data)
            st.success(f"Extraction complete — {len(df)} valid records found!")
            st.dataframe(df, use_container_width=True, hide_index=True)
            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Debit", f"{total_debit:,.2f}")
                col2.metric("Total Credit", f"{total_credit:,.2f}")
                col3.metric("Net", f"{net:,.2f}")
            except Exception as e:
                st.error(f"Totals error: {e}")

            st.download_button(
                "Download Excel",
                data=to_excel_bytes(data),
                file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No structured data detected. Check GPT response above.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
