import os, json, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# =============================================
# Load API key safely
# =============================================
try:
    from dotenv import load_dotenv
    load_dotenv()
except ModuleNotFoundError:
    st.warning("⚠️ 'python-dotenv' not installed — continuing without .env support.")

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No OpenAI API key found. Add it to .env or Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# =============================================
# Streamlit setup
# =============================================
st.set_page_config(page_title="📄 DataFalcon — Vendor Statement Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor")

# =============================================
# Helper functions
# =============================================
def extract_text_from_pdf(file):
    """Extract text from PDF pages."""
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())

def normalize_number(value):
    """Normalize EU/Spanish formatted numbers into float-compatible string."""
    if not value:
        return ""
    s = str(value).strip()

    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", s):  # EU 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^\d{1,3}(,\d{3})*\.\d{2}$", s):  # US 1,234.56
        s = s.replace(",", "")
    elif re.match(r"^\d+,\d{2}$", s):  # 150,00
        s = s.replace(",", ".")
    else:
        s = re.sub(r"[^\d.-]", "", s)
    return s

def extract_tax_id(raw_text):
    """Detect Spanish CIF/NIF or European VAT/AFM patterns in the raw text."""
    patterns = [
        r"\b[A-Z]{1}\d{7}[A-Z0-9]{1}\b",        # Spanish CIF/NIF (B12345678)
        r"\bES\d{9}\b",                         # Spanish VAT ES123456789
        r"\bEL\d{9}\b",                         # Greek VAT
        r"\b[A-Z]{2}\d{8,12}\b",                # Generic EU VAT
    ]
    for pat in patterns:
        match = re.search(pat, raw_text)
        if match:
            return match.group(0)
    return None

# =============================================
# Extraction logic — with all your conditions
# =============================================
def extract_with_llm(raw_text):
    """
    Extracts structured data from Spanish vendor statements.
    Applies logic for all prefixes, Debe/Haber structure, and payments/CN recognition.
    """
    prompt = f"""
    You are a professional accountant AI. 
    Read the following Spanish vendor statement and extract ONLY real documents (invoices or credit notes).

    For each document, return:
    - Alternative Document → the invoice/document number (look for prefixes: "Factura", "Documento", "Doc", "No", "Nº", "Num", "Número", "Nro", "Invoice")
    - Date → from "Fecha" or nearby text
    - Reason → "Invoice" or "Credit Note"
    - Document Value → numeric value (positive for invoice, negative for credit note)
    - Tax ID → CIF/NIF/VAT if present

    Rules for recognition:
    • Columns or text "Debe", "Debit", "Debe.", "Db", "Cargo", "Importe", "Valor", "Total", "Totale", "Amount" 
        → represent the document amount (Invoice → positive).
    • "Haber", "Credit", "Crédito", "Pago", "Transferencia", "Remesa", "Domiciliación" 
        → are payments → IGNORE unless text includes "Abono", "Nota de crédito", or "Devolución".
    • "Abono", "Nota de crédito", "Nota credito", "Devolución" → Credit Note → keep and mark negative.
    • If "Debe/Haber" columns not found, detect total amount near "Total", "Importe", "Valor", "Totale", "Amount".
    • Never duplicate documents.
    • Ignore lines that are purely payment, remittance, or bank operations.

    Output a VALID JSON array with fields:
    ["Alternative Document", "Date", "Reason", "Document Value", "Tax ID"]

    Example:
    [
      {{
        "Alternative Document": "2024/009",
        "Date": "12/09/2024",
        "Reason": "Invoice",
        "Document Value": "450.00",
        "Tax ID": "B12345678"
      }},
      {{
        "Alternative Document": "NC-102",
        "Date": "30/09/2024",
        "Reason": "Credit Note",
        "Document Value": "-150.00",
        "Tax ID": "B12345678"
      }}
    ]

    Text:
    \"\"\"{raw_text[:12000]}\"\"\"
    """

    response = client.responses.create(model=MODEL, input=prompt)
    content = response.output_text.strip()

    try:
        json_match = re.search(r'\[.*\]', content, re.DOTALL)
        if json_match:
            content = json_match.group(0)
        data = json.loads(content)
    except Exception as e:
        st.error(f"⚠️ Could not parse GPT output: {e}")
        st.text_area("🔍 Raw GPT Output", content[:2000], height=200)
        return []

    # --- Post-processing logic ---
    tax_id = extract_tax_id(raw_text)
    filtered = []

    for row in data:
        reason = str(row.get("Reason", "")).lower()
        val = normalize_number(row.get("Document Value", ""))

        if not val:
            continue

        # Skip payments entirely
        if any(w in reason for w in ["pago", "transferencia", "remesa", "domiciliacion"]):
            continue

        try:
            amount = float(val)
        except:
            amount = 0.0

        # Detect credit notes
        if any(w in reason for w in ["abono", "nota de crédito", "nota credito", "devolucion"]):
            amount = -abs(amount)
            row["Reason"] = "Credit Note"
        elif any(w in reason for w in ["factura", "invoice", "servicio", "mantenimiento", "documento", "doc"]):
            row["Reason"] = "Invoice"
        else:
            row["Reason"] = "Invoice" if amount > 0 else "Credit Note"

        row["Document Value"] = f"{amount:.2f}"
        row["Tax ID"] = tax_id if tax_id else "Missing TAX ID"
        filtered.append(row)

    return filtered

# =============================================
# Excel output helper
# =============================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

# =============================================
# Streamlit interface
# =============================================
uploaded_pdf = st.file_uploader("📂 Upload vendor statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded_pdf))

    st.text_area("🔍 Extracted text preview", text[:2000], height=200)

    if st.button("🤖 Extract data to Excel"):
        with st.spinner("Analyzing with GPT... please wait..."):
            data = extract_with_llm(text)

        if data:
            df = pd.DataFrame(data)
            st.success("✅ Extraction complete — invoices & credit notes detected successfully!")
            st.dataframe(df, use_container_width=True)

            excel_bytes = to_excel_bytes(data)
            st.download_button(
                "⬇️ Download Excel (Vendor Statement)",
                data=excel_bytes,
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No valid document data found. Verify your PDF content or layout.")
