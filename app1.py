import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Hybrid Vendor Statement Extractor")

# Load API key
try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No OpenAI API key found. Add it to .env or Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(value):
    """Normalize decimals like 1.234,56 → 1234.56."""
    if not value:
        return ""
    s = str(value).strip().replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return round(float(s), 2)
    except:
        return ""

def extract_raw_lines(uploaded_pdf):
    """Extract readable text lines from PDF (even if not tabular)."""
    lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", line):
                    lines.append(line.strip())
    return lines

# ==========================================================
# GPT EXTRACTOR
# ==========================================================
def extract_with_gpt(lines):
    """Send extracted lines to GPT for structured parsing."""
    joined_text = "\n".join(lines[:200])  # limit for safety

    prompt = f"""
You are an expert Spanish accountant.

Below are text lines from a vendor statement.
Each line may contain several numbers — typically a DEBE value (second-to-last) and a SALDO (last).
Your job:
1. Extract all invoice or credit note lines.
2. For each, return:
   - "Alternative Document": invoice or reference number (like 6--483)
   - "Date": dd/mm/yy or dd/mm/yyyy
   - "Reason": "Invoice" or "Credit Note"
   - "Document Value": the second-to-last numeric value (DEBE)
     - If the line mentions ABONO or NOTA DE CRÉDITO, make the value negative.
3. Ignore SALDO and totals.
4. Return valid JSON array only.

Lines:
\"\"\"{joined_text}\"\"\"
"""

    try:
        response = client.responses.create(model=MODEL, input=prompt)
        content = response.output_text.strip()
        json_match = re.search(r"\[.*\]", content, re.DOTALL)
        json_text = json_match.group(0) if json_match else content
        data = json.loads(json_text)
    except Exception as e:
        st.error(f"⚠️ GPT extraction failed: {e}")
        st.text_area("🔍 Raw GPT Output", content[:2000], height=200)
        return []

    # Post-clean
    cleaned = []
    for row in data:
        val = normalize_number(row.get("Document Value"))
        if val == "":
            continue
        reason = row.get("Reason", "").lower()
        if "credit" in reason or "abono" in reason or "nota de crédito" in reason:
            val = -abs(val)
            reason = "Credit Note"
        else:
            reason = "Invoice"
        cleaned.append({
            "Alternative Document": row.get("Alternative Document", "").strip(),
            "Date": row.get("Date", "").strip(),
            "Reason": reason,
            "Document Value": val
        })
    return cleaned

def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT APP
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        lines = extract_raw_lines(uploaded_pdf)

    if not lines:
        st.warning("⚠️ No readable text lines found. Check if the PDF is scanned.")
    else:
        st.text_area("📄 Extracted Sample (first 25 lines):", "\n".join(lines[:25]), height=250)

        if st.button("🤖 Analyze with GPT-4o-mini"):
            with st.spinner("Analyzing data... please wait..."):
                data = extract_with_gpt(lines)

            if not data:
                st.warning("⚠️ No structured invoice data detected.")
            else:
                df = pd.DataFrame(data)
                st.success(f"✅ Extraction complete — {len(df)} records found (Hybrid Mode).")
                st.dataframe(df, use_container_width=True)
                st.download_button(
                    "⬇️ Download Excel",
                    data=to_excel_bytes(data),
                    file_name="vendor_statement_hybrid.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("Please upload a vendor statement PDF to begin.")
