import os, json, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# ========== Load key ==========
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

# ========== UI ==========
st.set_page_config(page_title="📄 Vendor Statement Extractor", layout="wide")
st.title("📄 Vendor Statement → Excel Extractor (High Precision)")

# ========== Helpers ==========
def extract_text_from_pdf(file):
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for p in doc:
            text += p.get_text("text") + "\n"
    return text

def clean_text(t):
    return " ".join(t.replace("\xa0", " ").replace("€", " EUR").split())

def extract_with_llm(raw_text):
    """
    Ask GPT to extract clean JSON with strict financial field mapping.
    Debe=Debit, Haber=Credit, keep numeric only, use dot as decimal.
    """
    prompt = f"""
You are a meticulous accountant AI. 
Read the Spanish vendor statement and extract all invoice lines.

Rules:
- Debe = Debit
- Haber = Credit
- Never put both values in the same column.
- Use empty string if value missing.
- Convert 1.234,56 or 1,234.56 → 1234.56
- Return ONLY valid JSON array:
[
  {{"Invoice_Number":"...", "Date":"...", "Description":"...", "Debit":"...", "Credit":"...", "Balance":"..."}}
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

    # --- sanity check corrections ---
    for r in data:
        # normalize commas/periods
        for f in ["Debit", "Credit", "Balance"]:
            v = str(r.get(f, "")).replace(",", ".").replace(" ", "")
            if v and not re.match(r"^[0-9.]+$", v):
                v = re.sub(r"[^\d.]", "", v)
            r[f] = v
        # move numeric from Debit to Credit if both filled
        d, c = r.get("Debit", ""), r.get("Credit", "")
        if d and c:
            # choose smaller one as credit if typical pattern
            try:
                if float(c) < float(d):
                    r["Credit"], r["Debit"] = c, ""
                else:
                    r["Debit"], r["Credit"] = d, ""
            except:
                pass
    return data

def to_excel_bytes(records):
    df = pd.DataFrame(records)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

# ========== Streamlit logic ==========
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
            st.success("✅ Extraction complete (high precision)!")
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
