# =============================================
# app1.py — Vendor Statement → Excel Extractor
# =============================================

import os
import json
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# =============================================
# 1️⃣  Try to load .env (optional)
# =============================================
try:
    from dotenv import load_dotenv
    load_dotenv()  # Load .env if exists
except ModuleNotFoundError:
    st.warning("⚠️ 'python-dotenv' not installed — continuing without .env support.")

# =============================================
# 2️⃣  Secure API key loading (supports local + Streamlit Cloud)
# =============================================
api_key = (
    os.getenv("OPENAI_API_KEY")
    or st.secrets.get("OPENAI_API_KEY", None)
)

if not api_key:
    st.error(
        "❌ No OpenAI API key found.\n\n"
        "Please add it in one of these ways:\n"
        "1️⃣  Create a `.env` file with line: `OPENAI_API_KEY=your_key_here`\n"
        "2️⃣  Or, in Streamlit Cloud → Settings → Secrets → add the same line."
    )
    st.stop()

# Initialize OpenAI client
client = OpenAI(api_key=api_key)
MODEL = "gpt-4.1-mini"

# =============================================
# 3️⃣  Streamlit setup
# =============================================
st.set_page_config(page_title="📄 Vendor Statement Extractor", layout="wide")
st.title("📄 Vendor Statement → Excel Extractor (Spanish PDFs)")

# =============================================
# 4️⃣  Helper functions
# =============================================
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

# =============================================
# 5️⃣  Streamlit interface
# =============================================
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
