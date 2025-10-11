import os, json, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# =============================================
# Load environment variables safely
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
st.set_page_config(page_title="📄 Vendor Statement Extractor", layout="wide")
st.title("🦅 DataFalcon")

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
    """Normalize Spanish/EU formatted numbers like 1.234,56 or 1,234.56 into 1234.56"""
    if not value:
        return ""
    s = str(value).strip()

    # Case 1: European format 1.234,56 → 1234.56
    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    # Case 2: US format 1,234.56 → 1234.56
    elif re.match(r"^\d{1,3}(,\d{3})*\.\d{2}$", s):
        s = s.replace(",", "")
    # Case 3: Simple 150,00 → 150.00
    elif re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")
    # Case 4: Already normalized or other → clean stray symbols
    else:
        s = re.sub(r"[^\d.]", "", s)

    return s

def extract_with_llm(raw_text):
    """Send cleaned text to GPT and return structured JSON with correct columns."""
    prompt = f"""
    You are an expert accountant AI.

    Extract all invoice lines from the following Spanish vendor statement.
    Each line has: Invoice_Number, Date, Description, Debit (Debe), Credit (Haber), Balance (Saldo).

    Rules:
    - "Debe" → Debit column.
    - "Haber" → Credit column.
    - Words like "Pago" or "Abono" mean Credit.
    - Always include Balance as the rightmost value in each row.
    - Only one of Debit or Credit can have a value.
    - Return valid JSON array only.

    Example:
    [
      {{
        "Invoice_Number": "2025.TPY.190.1856",
        "Date": "12/09/2025",
        "Description": "Factura de servicios",
        "Debit": "3250.00",
        "Credit": "",
        "Balance": "3250.00"
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

    # --- Post-correction logic ---
    for row in data:
        for f in ["Debit", "Credit", "Balance"]:
            row[f] = normalize_number(row.get(f, ""))
        desc = row.get("Description", "").lower()
        # Ensure "Pago" or "Abono" entries are Credit
        if "pago" in desc or "abono" in desc:
            if row.get("Debit") and not row.get("Credit"):
                row["Credit"], row["Debit"] = row["Debit"], ""
        # Ensure only one side has a value
        if row.get("Debit") and row.get("Credit"):
            try:
                d, c = float(row["Debit"]), float(row["Credit"])
                if "pago" in desc or "abono" in desc or c < d:
                    row["Credit"], row["Debit"] = c, ""
                else:
                    row["Debit"], row["Credit"] = d, ""
            except:
                pass
    return data

def to_excel_bytes(records):
    df = pd.DataFrame(records)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

# =============================================
# Streamlit interface
# =============================================
uploaded_pdf = st.file_uploader("📂 Upload a vendor statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded_pdf))

    st.text_area("🔍 Extracted text preview", text[:2000], height=200)

    if st.button("🤖 Extract data to Excel"):
        with st.spinner("Analyzing with GPT... please wait..."):
            data = extract_with_llm(text)

        if data:
            df = pd.DataFrame(data)
            st.success("✅ Extraction complete (decimal fix)!")
            st.dataframe(df)

            excel_bytes = to_excel_bytes(data)
            st.download_button(
                "⬇️ Download Excel",
                data=excel_bytes,
                file_name="statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No structured data found. Try another PDF or verify text extraction.")
