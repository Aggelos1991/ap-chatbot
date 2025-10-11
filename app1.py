# =============================================
# app1.py — Vendor Statement → Excel Extractor
# =============================================

import os
import json
import re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# =============================================
# 1️⃣  Load environment variables safely
# =============================================
try:
    from dotenv import load_dotenv
    load_dotenv()
except ModuleNotFoundError:
    st.warning("⚠️ 'python-dotenv' not installed — continuing without .env support.")

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error(
        "❌ No OpenAI API key found.\n\n"
        "Please add it in one of these ways:\n"
        "1️⃣ Create a `.env` file with line: OPENAI_API_KEY=your_key_here\n"
        "2️⃣ Or, in Streamlit Cloud → Settings → Secrets → add the same line."
    )
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"  # stable and fast for structured extraction

# =============================================
# 2️⃣  Streamlit setup
# =============================================
st.set_page_config(page_title="📄 Vendor Statement Extractor", layout="wide")
st.title("📄 Vendor Statement → Excel Extractor (Spanish PDFs)")

# =============================================
# 3️⃣  Helper functions
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
    """Send text to GPT and return structured JSON with fallbacks."""
    prompt = f"""
    You are an expert in financial data extraction.
    Extract all invoice entries from the following Spanish vendor statement.

    Return ONLY a valid JSON array like:
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

    # Try multiple fallback parsing strategies
    try:
        # Extract JSON between brackets if GPT adds text around
        json_match = re.search(r'\[.*\]', content, re.DOTALL)
        if json_match:
            content = json_match.group(0)
        data = json.loads(content)
    except Exception as e:
        st.error(f"⚠️ Could not parse GPT output: {e}")
        st.text_area("🔍 Raw GPT Output", content[:2000], height=200)
        data = []
    return data


def to_excel_bytes(records):
    """Return Excel file in memory."""
    df = pd.DataFrame(records)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

# =============================================
# 4️⃣  Streamlit interface
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
