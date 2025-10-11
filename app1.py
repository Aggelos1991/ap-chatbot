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
st.title("📄 Vendor Statement → Excel Extractor (Full Precision + Balance)")

# =============================================
# Helper functions
# =============================================
def extract_text_from_pdf(file):
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())

def parse_number(value):
    """Normalize numeric strings to consistent float-like strings."""
    if not value:
        return ""
    value = str(value).replace(",", ".")
    value = re.sub(r"[^\d.]", "", value)
    return value

def extract_with_llm(raw_text):
    """Send text to GPT and return structured JSON with all columns preserved."""
    prompt = f"""
You are a precise financial data extractor.

Read this Spanish vendor statement and extract every invoice line into JSON with exactly:
Invoice_Number, Date, Description, Debit (Debe), Credit (Haber), Balance (Saldo).

Rules:
- If "Debe" (Debit) is filled, "Haber" (Credit) must be empty, and vice versa.
- Always include "Balance" exactly as shown in the statement.
- Return numeric values as strings using dot as decimal separator.
- Do NOT omit or skip the Balance column.
- Return ONLY valid JSON array, no text before or after.

Example output:
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

    # --- post-correction layer ---
    for row in data:
        for field in ["Debit", "Credit", "Balance"]:
            row[field] = parse_number(row.get(field, ""))
        # If Debit and Credit both exist, fix based on smaller/larger logic
        d, c = row.get("Debit", ""), row.get("Credit", "")
        if d and c:
            try:
                if float(c) < float(d):
                    row["Credit"], row["Debit"] = c, ""
                else:
                    row["Debit"], row["Credit"] = d, ""
            except:
                pass
        # Ensure balance never disappears
        if not row.get("Balance"):
            row["Balance"] = row.get("Debit") or row.get("Credit") or ""
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
            st.success("✅ Extraction complete (Balance restored)!")
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
