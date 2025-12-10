import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ==========================================================
# STREAMLIT CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — ULTRA FAST VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — Ultra Fast SABA Extractor")

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No API key")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# PDF → RAW LINES
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
                if not line.strip():
                    continue
                if "Saldo" in line or "SALDO" in line:
                    continue
                lines.append(line)
    return lines

# ==========================================================
# GPT FAST EXTRACTOR
# ==========================================================
def extract_with_gpt_fast(lines):
    all_rows = []
    BATCH = 20  # 🚀 SUPER FAST

    for i in range(0, len(lines), BATCH):
        block = "\n".join(lines[i:i+BATCH])

        prompt = f"""
Extract rows from SABA ledger lines.

RULES:
- First dd/mm/yyyy = Date
- Last two numeric values = Debit, Credit
- If Debit negative → Credit Note
- Alternative Document = first long number/code after date
- Return ONLY JSON array.

LINES:
{block}
"""

        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            content = response.choices[0].message.content.strip()

            # Force JSON extraction
            json_array = re.search(r"\[.*\]", content, re.DOTALL)
            if not json_array:
                continue

            rows = json.loads(json_array.group(0))
            all_rows.extend(rows)

        except Exception as e:
            st.warning(f"Batch error: {e}")
            continue

    return all_rows

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
pdf_file = st.file_uploader("📂 Upload Ledger PDF", type=["pdf"])

if pdf_file:
    with st.spinner("Extracting lines…"):
        lines = extract_lines(pdf_file)

    st.success(f"Extracted {len(lines)} lines.")
    st.text_area("Preview:", "\n".join(lines[:30]), height=300)

    if st.button("⚡ Run Ultra-Fast Extraction", type="primary"):
        with st.spinner("GPT extracting… (fast mode)"):
            data = extract_with_gpt_fast(lines)

        if not data:
            st.error("No rows extracted.")
        else:
            df = pd.DataFrame(data)
            st.success(f"Done! {len(df)} rows extracted.")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel(data),
                file_name="saba_ledger.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
