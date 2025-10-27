import os
import json
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# ==========================
# STREAMLIT CONFIG
# ==========================
st.set_page_config(page_title="📄 Vendor Statement Extractor", layout="wide")
st.title("📄 Vendor Statement → Excel Extractor (Spanish PDFs)")

# ==========================
# LOAD OPENAI API KEY SAFELY
# ==========================
API_KEY = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")

if not API_KEY:
    st.error("❌ OpenAI API key not found. Please set it as an environment variable or in Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=API_KEY)
MODEL = "gpt-4.1-mini"

# ==========================
# HELPER FUNCTIONS
# ==========================
def extract_text_from_pdf(file):
    """Extract text from all PDF pages."""
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    """Normalize spaces and symbols."""
    text = text.replace("\xa0", " ").replace("€", " EUR")
    text = " ".join(text.split())
    return text

def extract_with_llm(raw_text):
    """Send text to GPT and return structured JSON."""
    prompt = f"""
    From the following Spanish vendor statement, extract each invoice line
    with these fields:
    - Invoice_Number
    - Date
    - Description
    - Debit (Debe)
    - Credit (Haber)
    - Balance (Saldo)
    Return ONLY valid JSON array.
    Text:
    \"\"\"{raw_text[:12000]}\"\"\"
    """
    response = client.responses.create(model=MODEL, input=prompt)
    content = response.output_text.strip()

    try:
        data = json.loads(content)
    except Exception:
        # Fallback in case GPT adds extra markdown fences
        content = content.split("```")[-1]
        data = json.loads(content)
    return data

def to_excel_bytes(records):
    """Return Excel file in memory."""
    df = pd.DataFrame(records)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

# ==========================
# STREAMLIT INTERFACE
# ==========================
uploaded_pdf = st.file_uploader("📂 Upload a vendor statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        text = extract_text_from_pdf(uploaded_pdf)
        cleaned = clean_text(text)

    st.text_area("🔍 Extracted text preview", cleaned[:2000], height=200)

    if st.button("🤖 Extract data to Excel"):
        with st.spinner("Analyzing with GPT... please wait..."):
            try:
                data = extract_with_llm(cleaned)
            except Exception as e:
                st.error(f"⚠️ LLM extraction failed: {e}")
                st.stop()

        if data:
            df = pd.DataFrame(data)
            st.success("✅ Extraction complete!")
            st.dataframe(df)

            excel_bytes = to_excel_bytes(data)
            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name="statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No structured data found. Try another PDF or verify text extraction.")
