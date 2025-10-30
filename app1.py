import os
import json
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image
import streamlit as st
from openai import OpenAI

# ==========================
# STREAMLIT CONFIG
# ==========================
st.set_page_config(page_title="📄 Vendor Statement Extractor (OCR Ready)", layout="wide")
st.title("📄 Vendor Statement Extractor (with OCR fallback)")

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
    """Extract text from PDF, with OCR fallback for scanned pages."""
    text = ""
    ocr_pages = 0

    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page_number, page in enumerate(doc, start=1):
            page_text = page.get_text("text")

            # If no text found (scanned image), run OCR
            if not page_text.strip():
                pix = page.get_pixmap(dpi=300)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                ocr_text = pytesseract.image_to_string(img, lang="spa")  # OCR in Spanish
                text += ocr_text + "\n"
                ocr_pages += 1
            else:
                text += page_text + "\n"

    if ocr_pages > 0:
        st.warning(f"⚙️ {ocr_pages} page(s) processed via OCR (image-based PDF).")

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
    with st.spinner("📄 Extracting text from PDF (auto OCR if needed)..."):
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
else:
    st.info("Please upload a vendor statement PDF to begin.")
