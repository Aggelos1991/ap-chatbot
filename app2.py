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
st.set_page_config(page_title="Vendor Statement Extractor", layout="wide")
st.title("Vendor Statement → Excel Extractor (Spanish PDFs)")

# ==========================
# LOAD OPENAI API KEY SAFELY
# ==========================
API_KEY = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not API_KEY:
    st.error("OpenAI API key not found. Set it as an environment variable or in Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=API_KEY)
MODEL = "gpt-4o-mini"  # Updated to latest recommended model

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
    text = text.replace("\xa0", " ").replace("€", " EUR").replace("'", " ")
    text = " ".join(text.split())
    return text


def extract_with_llm(raw_text):
    """Send text to GPT and return a clean JSON array."""
    prompt = f"""
    You are an expert at extracting invoice lines from Spanish vendor statements.
    Return **ONLY** a valid JSON array of objects with these exact keys (no extra text, no markdown):

    - Invoice_Number (string)
    - Date          (string, format DD/MM/YYYY or YYYY-MM-DD)
    - Description   (string)
    - Debit         (number, use 0 if empty or missing)
    - Credit        (number, use 0 if empty or missing)
    - Balance       (number, use 0 if empty or missing)

    Text to analyze (first 12,000 characters):
    \"\"\"{raw_text[:12000]}\"\"\"
    """

    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": "You return ONLY valid JSON. No explanations. No markdown."},
                {"role": "user", "content": prompt}
            ],
            temperature=0,
            max_tokens=1500
        )
        content = response.choices[0].message.content.strip()

        # --- Robust JSON extraction ---
        if "```" in content:
            parts = content.split("```")
            json_part = parts[1] if len(parts) > 1 else parts[0]
            if json_part.lower().startswith("json"):
                json_part = json_part[4:].strip()
            content = json_part.strip()

        data = json.loads(content)
        return data

    except json.JSONDecodeError as e:
        st.error(f"JSON parsing failed: {e}")
        st.code(content, language="json")
        st.stop()
    except Exception as e:
        st.error(f"LLM call failed: {e}")
        st.stop()


def to_excel_bytes(records):
    """Return Excel file in memory."""
    df = pd.DataFrame(records)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Statement')
    output.seek(0)
    return output


# ==========================
# STREAMLIT INTERFACE
# ==========================
st.markdown("### Upload a Spanish vendor statement (PDF) to extract invoice lines into Excel")

uploaded_pdf = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_pdf:
    with st.spinner("Extracting text from PDF..."):
        text = extract_text_from_pdf(uploaded_pdf)
        cleaned = clean_text(text)

    st.text_area("Extracted Text Preview", cleaned[:2000], height=200)

    if st.button("Extract Data to Excel", type="primary"):
        with st.spinner("Analyzing with GPT... this may take 10-20 seconds"):
            try:
                data = extract_with_llm(cleaned)
            except Exception:
                st.stop()

        if not data:
            st.warning("No data extracted. The PDF may not contain tabular invoice lines.")
            st.stop()

        df = pd.DataFrame(data)

        # Clean numeric columns
        for col in ['Debit', 'Credit', 'Balance']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        st.success("Extraction complete!")
        st.dataframe(df, use_container_width=True)

        excel_bytes = to_excel_bytes(data)
        st.download_button(
            label="Download Excel File",
            data=excel_bytes,
            file_name="vendor_statement_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info("Tip: If results are incomplete, try a clearer PDF or split large files.")
