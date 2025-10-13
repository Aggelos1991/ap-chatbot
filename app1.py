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
    pass

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
st.title("🦅 DataFalcon — Vendor Statement Extractor (Optimized DEBE Edition)")

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


def clean_text(text: str) -> str:
    """Normalize whitespace and characters."""
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())


def normalize_number(value):
    """Normalize European / US number formats to float-compatible string."""
    if not value:
        return ""
    s = str(value).strip()
    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", s):  # 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^\d+,\d{2}$", s):  # 150,00
        s = s.replace(",", ".")
    elif re.match(r"^\d{1,3}(,\d{3})*\.\d{2}$", s):  # 1,234.56
        s = s.replace(",", "")
    else:
        s = re.sub(r"[^\d.-]", "", s)
    return s


def extract_tax_id(raw_text):
    """Detect Spanish / EU VAT ID."""
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
# CORE EXTRACTION
# =============================================
def extract_with_llm(raw_text):
    """
    Use GPT to identify structured invoice lines, then
    post-process to ensure only DEBE (first numeric value) is kept.
    """
    prompt = f"""
    You are an expert Spanish accountant AI.
    Extract all invoice or credit note lines from this vendor statement.

    Each record must include:
      - Invoice_Number (Factura / Documento / Nº / Num / Número / Doc)
      - Date (Fecha)
      - Description (Concepto)
      - Numeric values (Debe / Haber / Saldo)

    Rules:
      - "Debe" = invoice value (we care about this)
      - Ignore "Haber", "Saldo", "Balance", "Pago", "Banco", "Remesa", "Cobro".
      - If multiple numbers exist, the first numeric value is the DEBE.
      - Return only valid JSON.

    Text:
    \"\"\"{raw_text[:12000]}\"\"\"
    """

    try:
        response = client.responses.create(model=MODEL, input=prompt)
        content = response.output_text.strip()
        json_match = re.search(r"\[.*\]", content, re.DOTALL)
        content = json_match.group(0) if json_match else content
        data = json.loads(content)
    except Exception as e:
        st.error(f"⚠️ Could not parse GPT output: {e}")
        st.text_area("🔍 Raw GPT Output", content[:2000], height=200)
        return []

    cleaned = []
    for row in data:
        # Collect all numeric values found in row text
        values = []
        for key in row:
            val = normalize_number(row.get(key, ""))
            if re.match(r"^\d+(\.\d+)?$", val):
                values.append(val)

        # Always take the first numeric as DEBE
        debe = values[0] if values else ""
        row_clean = {
            "Alternative Document": row.get("Invoice_Number", ""),
            "Date": row.get("Date", ""),
            "Reason": "Invoice",
            "Document Value": debe,
        }
        cleaned.append(row_clean)

    return cleaned

# =============================================
# EXPORT
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

    st.text_area("🔍 Extracted Text Preview", text[:2000], height=200)

    if st.button("🤖 Extract Data to Excel"):
        with st.spinner("Analyzing with GPT... please wait..."):
            data = extract_with_llm(text)

        tax_id = extract_tax_id(text)
        for row in data:
            row["Tax ID"] = tax_id if tax_id else "Missing TAX ID"

        if data:
            df = pd.DataFrame(data)
            st.success("✅ Extraction complete — DEBE values only (Saldo/Haber ignored).")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No valid invoice data found. Try another PDF or verify formatting.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
