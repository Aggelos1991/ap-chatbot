import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ==========================================================
# STREAMLIT CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Final Version", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL VERSION")

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No API key")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"   # fast + cheap + accurate

# ==========================================================
# RAW LINE EXTRACTION
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
                if not line:
                    continue
                if "Saldo" in line or "SALDO" in line:
                    continue
                lines.append(line)
    return lines


# ==========================================================
# GPT EXTRACTOR — ULTRA FAST + STABLE
# ==========================================================
def extract_with_gpt(lines):
    final_rows = []
    BATCH = 20  # SUPER-fast, SUPER-stable

    for i in range(0, len(lines), BATCH):
        block = "\n".join(lines[i:i+BATCH])

        prompt = f"""
Extract structured accounting rows from SABA ledger lines.

RULES:
- First dd/mm/yyyy = Date
- Last two numeric values = Debit, Credit
- If Debit < 0 → treat as Credit Note (move to credit)
- Alternative Document = FIRST long number/code after date
- Always return clean JSON array.

LINES:
{block}
"""

        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            content = response.choices[0].message.content

            # Extract JSON array
            match = re.search(r"\[.*\]", content, re.DOTALL)
            if not match:
                continue
            
            rows = json.loads(match.group(0))
            final_rows.extend(rows)

        except Exception as e:
            st.warning(f"Batch error: {e}")
            continue

    return final_rows


# ==========================================================
# EXPORTER (FIXED PYARROW ISSUE)
# ==========================================================
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
    with st.spinner("Extracting PDF lines…"):
        lines = extract_lines(pdf_file)

    st.success(f"Extracted {len(lines)} lines.")
    st.text_area("Preview:", "\n".join(lines[:30]), height=300)

    if st.button("⚡ Run Extraction"):
        with st.spinner("GPT extracting…"):
            data = extract_with_gpt(lines)

        if not data:
            st.error("No rows extracted.")
            st.stop()

        df = pd.DataFrame(data)

        # =======================================================
        # FIX: Convert Debit/Credit to REAL FLOATS (no PyArrow crash)
        # =======================================================
        def normalize(x):
            if x is None: return None
            x = str(x)
            x = x.replace(".", "").replace(",", ".")
            x = re.sub(r"[^\d\.-]", "", x)
            try:
                return float(x)
            except:
                return None

        if "Debit" in df.columns:
            df["Debit"] = df["Debit"].apply(normalize)

        if "Credit" in df.columns:
            df["Credit"] = df["Credit"].apply(normalize)

        st.success(f"Done! {len(df)} rows extracted.")

        st.dataframe(df, use_container_width=True)

        st.download_button(
            "⬇️ Download Excel",
            data=to_excel(df),
            file_name="ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
