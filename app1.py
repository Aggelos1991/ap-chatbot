import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ==========================================================
# STREAMLIT CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — FINAL VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL SABA LEDGER VERSION")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ Missing API key")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# CLEAN NUMBER
# ==========================================================
def clean_num(v):
    if not v:
        return ""
    v = v.replace(".", "").replace(",", ".")
    try:
        return round(float(v), 2)
    except:
        return ""

# ==========================================================
# EXTRACT RAW LINES FROM PDF
# ==========================================================
def extract_lines(pdf):
    lines = []
    with pdfplumber.open(pdf) as pdf_obj:
        for page in pdf_obj.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                line = " ".join(line.split())
                if (
                    not line.strip()
                    or "Saldo" in line
                    or "saldo" in line
                    or "SALDO" in line
                ):
                    continue
                lines.append(line)
    return lines

# ==========================================================
# GPT EXTRACTION — FINAL VERSION
# ==========================================================
def extract_with_gpt(lines):
    all_records = []
    BATCH = 60

    for i in range(0, len(lines), BATCH):
        batch = lines[i:i+BATCH]
        text_block = "\n".join(batch)

        prompt = f"""
You extract accounting ledger rows from SABA-style Spanish PDFs.

VERY IMPORTANT RULES:
1. The first date dd/mm/yyyy is the Date.
2. The **LAST TWO** numeric values are ALWAYS:
   - Debit (DEBE)
   - Credit (HABER)
3. A negative value means Credit Note → put in Credit.
4. "Alternative Document" must be the FIRST long number or code after the date.
5. SALDO is ignored.
6. Output STRICT JSON array. NO text, NO comments.

Example:
Line: "02/01/2023 GRL ... 840,95 6.648,01"
→ Debit: 840,95
→ Credit: 6648,01

Extract fields:
- Alternative Document
- Concepto (the whole text except the numbers)
- Date
- Reason: "Invoice" (if Debit), "Payment" (if Credit), "Credit Note" (if Credit was negative)
- Debit
- Credit

Return ONLY valid JSON array.

TEXT TO PARSE:
{text_block}
"""

        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            content = response.choices[0].message.content.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group(0))
                all_records.extend(data)
        except Exception as e:
            st.warning(f"GPT error: {e}")

    # clean numbers
    cleaned = []
    for r in all_records:
        r["Debit"] = clean_num(str(r.get("Debit", "")))
        r["Credit"] = clean_num(str(r.get("Credit", "")))
        cleaned.append(r)

    return cleaned

# ==========================================================
# EXPORT
# ==========================================================
def to_excel(data):
    df = pd.DataFrame(data)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# UI
# ==========================================================
pdf_file = st.file_uploader("📂 Upload SABA Ledger PDF", type=["pdf"])

if pdf_file:
    with st.spinner("Extracting lines from PDF…"):
        lines = extract_lines(pdf_file)

    st.success(f"Extracted {len(lines)} lines.")
    st.text_area("Preview first 30 lines:", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Final Extraction", type="primary"):
        with st.spinner("GPT analyzing…"):
            data = extract_with_gpt(lines)

        if not data:
            st.error("No valid rows detected.")
        else:
            df = pd.DataFrame(data)
