import os
import json
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# ==========================
# CONFIG
# ==========================
st.set_page_config(page_title="Vendor Statement Extractor", layout="wide")
st.title("Vendor Statement → Excel (Spanish PDFs)")

API_KEY = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not API_KEY:
    st.error("OpenAI API key missing.")
    st.stop()

client = OpenAI(api_key=API_KEY)
MODEL = "gpt-4o-mini"

# ==========================
# HELPERS
# ==========================
def extract_text_from_pdf(file):
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())

def extract_with_llm(raw_text):
    prompt = f"""
    Extract every invoice line from the Spanish vendor statement.
    Return **ONLY** a JSON array of objects with these exact keys:
    - Invoice_Number (string)
    - Date          (string, DD/MM/YYYY or YYYY-MM-DD)
    - Description   (string)
    - Debit         (number, 0 if empty)
    - Credit        (number, 0 if empty)
    - Balance       (number, 0 if empty)

    Text (max 12 000 chars):
    \"\"\"{raw_text[:12000]}\"\"\"
    """

    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": "Return ONLY valid JSON. No markdown, no explanations."},
            {"role": "user",   "content": prompt}
        ],
        temperature=0,
        max_tokens=1500
    )
    json_str = resp.choices[0].message.content.strip()

    # strip possible markdown
    if "```" in json_str:
        json_str = json_str.split("```")[1].replace("json", "", 1).strip()

    return json.loads(json_str)

def to_excel_bytes(records):
    output = BytesIO()
    pd.DataFrame(records).to_excel(output, index=False)
    output.seek(0)
    return output

# ==========================
# UI
# ==========================
uploaded = st.file_uploader("Upload PDF", type="pdf")

if uploaded:
    with st.spinner("Reading PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded))

    st.text_area("Text preview", text[:2000], height=150)

    if st.button("Extract to Excel"):
        with st.spinner("Calling GPT..."):
            data = extract_with_llm(text)

        # 1. Show raw JSON (exactly what you loved)
        st.subheader("Raw JSON from LLM")
        st.json(data, expanded=False)

        # 2. Show table
        df = pd.DataFrame(data)
        for c in ["Debit", "Credit", "Balance"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

        st.subheader("Extracted Table")
        st.dataframe(df, use_container_width=True)

        # 3. Download
        st.download_button(
            "Download Excel",
            data=to_excel_bytes(data),
            file_name="statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
