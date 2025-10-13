import os, json, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# =============================================
# ENVIRONMENT SETUP
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
# STREAMLIT CONFIG
# =============================================
st.set_page_config(page_title="🦅 DataFalcon — Vendor Statement Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor (Stable Version)")

# =============================================
# HELPERS
# =============================================
def extract_text_from_pdf(file):
    """Safely extract text from uploaded PDF."""
    file_bytes = file.getvalue()
    if not file_bytes:
        raise ValueError("Uploaded file is empty or unreadable.")

    text = ""
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text


def clean_text(text):
    """Clean extracted text."""
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())


def normalize_number(value):
    """Normalize European/US number formats, handle negatives."""
    if not value:
        return ""
    s = str(value).strip()
    is_negative = s.startswith('-') or 'HABER' in s.upper()  # Assume HABER implies credit (negative)
    s = re.sub(r"[^\d.,-]", "", s)  # Remove non-numeric except , . -
    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", s):  # 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^\d{1,3}(,\d{3})*\.\d{2}$", s):  # 1,234.56
        s = s.replace(",", "")
    elif re.match(r"^\d+,\d{2}$", s):  # 150,00
        s = s.replace(",", ".")
    else:
        s = re.sub(r"[^\d.-]", "", s)
    try:
        num = float(s)
        return -num if is_negative else num
    except ValueError:
        return ""


def extract_tax_id(raw_text):
    """Detect CIF/NIF/VAT from text."""
    patterns = [
        r"\b[A-Z]{1}\d{7}[A-Z0-9]{1}\b",
        r"\bES\d{9}\b",
        r"\bEL\d{9}\b",
        r"\b[A-Z]{2}\d{8,12}\b",
    ]
    for pat in patterns:
        match = re.search(pat, raw_text)
        if match:
            return match.group(0)
    return None


# =============================================
# CORE EXTRACTION (STABLE + FILTERING)
# =============================================
def extract_with_llm(raw_text):
    """
    Extract structured invoice data from Spanish vendor statement.
    - Handles DEBE for invoices (positive amounts).
    - Handles HABER only for credit notes (negative amounts), not payments.
    - Captures invoice numbers with Spanish prefixes (e.g., Factura, Nº Factura).
    - Strict filtering for payments/balances.
    """

    prompt = f"""
    You are an expert accountant AI.

    Extract all *invoice* or *credit note* lines from the following Spanish vendor statement.
    Each line must include:
    - Invoice Number: Full number with Spanish prefixes (e.g., "Factura 6--483", "Nº Documento ABC-123", "Num Factura 456")
    - Date: Fecha in DD/MM/YY or similar
    - Type: "Invoice" if it's a factura/debit, "Credit Note" if it's a nota de crédito/abono
    - Amount: From DEBE/IMPORTE/VALOR/TOTAL/TOTALE/AMOUNT for invoices (positive). From HABER only if it's a credit note (make negative, e.g., -123.45). Normalize to US format (e.g., 123.45).

    ⚠️ Rules:
    - ONLY include lines that are explicitly invoices (factura) or credit notes (nota de crédito, abono).
    - For credit notes, use HABER as amount (negative), but ONLY if the line mentions "Nota de Crédito", "Abono", or similar — NOT for payments.
    - Do NOT include any payments, credits that are not notes, or balances: Ignore SALDO, BALANCE, ACUMULADO, RESTANTE, HABER (unless credit note), CRÉDITO (unless credit note), PAGO, BANCO, REMESA, COBRO, DOMICILIACIÓN, "Cobro Efecto", "Banco Santander".
    - Ignore any line with payment-related terms.

    Output as JSON array like:
    [
      {{
        "Invoice Number": "Factura 6--483",
        "Date": "24/01/25",
        "Type": "Invoice",
        "Amount": 322.27
      }},
      {{
        "Invoice Number": "Nota de Crédito 789",
        "Date": "15/02/25",
        "Type": "Credit Note",
        "Amount": -150.00
      }}
    ]

    Text:
    \"\"\"{raw_text[:12000]}\"\"\"
    """

    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}]
        )
        content = response.choices[0].message.content.strip()
        # Improved parsing: Strip code blocks if present
        content = re.sub(r"^```json|```$", "", content).strip()
        json_match = re.search(r"\[.*\]", content, re.DOTALL)
        content = json_match.group(0) if json_match else content
        data = json.loads(content)
    except json.JSONDecodeError as e:
        st.error(f"⚠️ Could not parse JSON from GPT: {e}")
        st.text_area("🔍 Raw GPT Output", content[:2000], height=200)
        return []
    except Exception as e:
        st.error(f"⚠️ GPT API error: {e}")
        return []

    # Post-correction logic
    for row in data:
        # Normalize amount (handle legacy "Document Value" if GPT uses it)
        amt_key = "Amount" if "Amount" in row else "Document Value"
        if amt_key in row:
            row["Amount"] = normalize_number(row[amt_key])
            if amt_key != "Amount":
                del row[amt_key]
        # If GPT extracted separate Debit/Credit, merge into Amount
        if "Debit" in row:
            row["Amount"] = normalize_number(row["Debit"])
            del row["Debit"]
        if "Credit" in row:
            row["Amount"] = normalize_number(row["Credit"]) * -1  # Make negative
            del row["Credit"]
        # Remove unwanted
        if "Balance" in row:
            del row["Balance"]
        # Infer Type if missing
        if "Type" not in row or not row["Type"]:
            row["Type"] = "Credit Note" if row.get("Amount", 0) < 0 else "Invoice"

    # Filter out any zero/empty amounts or suspicious entries
    data = [row for row in data if row.get("Amount") and abs(row["Amount"]) > 0]

    return data


# =============================================
# EXPORT TO EXCEL
# =============================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# =============================================
# STREAMLIT UI
# =============================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        try:
            text = clean_text(extract_text_from_pdf(uploaded_pdf))
        except Exception as e:
            st.error(f"❌ Failed to read PDF: {e}")
            st.stop()

    st.text_area("🔍 Extracted Text (first 2000 chars)", text[:2000], height=200)

    if st.button("🤖 Extract Data to Excel"):
        with st.spinner("Analyzing with GPT..."):
            data = extract_with_llm(text)

        if data:
            tax_id = extract_tax_id(text)
            for row in data:
                row["Tax ID"] = tax_id if tax_id else "Missing TAX ID"

            df = pd.DataFrame(data)
            st.success("✅ Extraction complete!")
            st.dataframe(df, use_container_width=True)

            excel_bytes = to_excel_bytes(data)
            st.download_button(
                "⬇️ Download Excel (Vendor Statement)",
                data=excel_bytes,
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No valid document data found. Try with a different PDF or check formatting.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
