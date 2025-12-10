import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
import time

# ==========================================================
# STREAMLIT CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Bulletproof", layout="wide")
st.title("🦅 DataFalcon Pro — BULLETPROOF VERSION")

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No API key found.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# Load PDF lines
# ==========================================================
def extract_lines(pdf):
    lines = []
    with pdfplumber.open(pdf) as pdf_obj:
        for page in pdf_obj.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                line = " ".join(line.split())
                if not line:
                    continue
                if "Saldo" in line or "SALDO" in line:
                    continue
                lines.append(line)
    return lines

# ==========================================================
# GPT Extractor — BULLETPROOF
# ==========================================================
def extract_batch(batch_lines):

    prompt = f"""
Extract structured rows from SABA ledger lines.
RULES:
- First dd/mm/yyyy = Date
- Last 2 numbers = Debit, Credit
- If Debit < 0 → move to Credit and set Debit null
- Alternative Document = first long numeric code after date
Return STRICT JSON array ONLY.

LINES:
{chr(10).join(batch_lines)}
"""

    # Maximum 3 retries
    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )

            content = response.choices[0].message.content.strip()
            match = re.search(r"\[.*\]", content, re.DOTALL)
            if not match:
                continue

            return json.loads(match.group(0))

        except Exception as e:
            if attempt == 2:
                st.warning(f"Failed batch after 3 tries: {e}")
                return []
            time.sleep(1)

    return []

# ==========================================================
# RUN EXTRACTION
# ==========================================================
def normalize_number(x):
    if x is None:
        return None
    x = str(x).replace(".", "").replace(",", ".")
    x = re.sub(r"[^\d\.-]", "", x)
    try:
        return float(x)
    except:
        return None

def to_excel(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# UI
# ==========================================================
pdf_file = st.file_uploader("📂 Upload Ledger PDF", type=["pdf"])

if pdf_file:
    with st.spinner("Extracting PDF text..."):
        lines = extract_lines(pdf_file)

    st.success(f"Extracted {len(lines)} lines.")
    st.text_area("Preview:", "\n".join(lines[:30]), height=300)

    if st.button("⚡ Run Extraction"):
        BATCH = 8
        all_rows = []

        with st.spinner("Processing batches with GPT..."):
            for i in range(0, len(lines), BATCH):
                batch = lines[i:i+BATCH]
                rows = extract_batch(batch)
                all_rows.extend(rows)

        if not all_rows:
            st.error("No rows extracted. Something went wrong.")
            st.stop()

        df = pd.DataFrame(all_rows)

        if "Debit" in df.columns:
            df["Debit"] = df["Debit"].apply(normalize_number)
        if "Credit" in df.columns:
            df["Credit"] = df["Credit"].apply(normalize_number)

        st.success(f"Extraction complete — {len(df)} rows found!")
        st.dataframe(df, use_container_width=True)

        st.download_button(
            "⬇️ Download Excel",
            data=to_excel(df),
            file_name="ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
