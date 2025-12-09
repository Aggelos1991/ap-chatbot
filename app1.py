# ==========================================================
# 🦅 DataFalcon Pro v3 — HYBRID GPT + Ultra Fast (FINAL)
# ==========================================================

import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import re
from openai import OpenAI

# ==========================================================
# CONFIG
# ==========================================================

st.set_page_config(page_title="🦅 DataFalcon Pro v3 — Hybrid GPT", layout="wide")
st.title("🦅 DataFalcon Pro v3 — Hybrid GPT Edition (FINAL)")

# Load key
api_key = st.secrets.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)

GPT_MODEL = "gpt-4o-mini"   # stable + fast


# ==========================================================
# HELPERS
# ==========================================================

def clean_amount(v):
    """Normalize numbers — EU format support."""
    if v is None or v == "":
        return None
    s = str(v).strip()

    s = re.sub(r"[^\d,.\-]", "", s)

    # Convert EU format 1.234,56 → 1234.56
    if "," in s and "." in s:
        if s.find(".") < s.find(","):
            s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return None


def extract_table_pdfplumber(file_bytes):
    """Extract table using pdfplumber."""
    records = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                header = table[0]
                for row in table[1:]:
                    records.append(dict(zip(header, row)))

    if len(records) == 0:
        return pd.DataFrame()

    return pd.DataFrame(records)


def gpt_extract_table(text):
    """Use GPT to convert messy text to a structured table."""
    prompt = f"""
You are a data extraction model. Convert the following text into a clean table.
Output strictly as JSON list of objects with:
["Referencia", "Debit", "Credit"]

Text:
{text}
"""

    response = client.responses.create(
        model=GPT_MODEL,
        input=prompt,
        max_output_tokens=2000
    )

    content = response.output_text
    try:
        df = pd.read_json(BytesIO(content.encode()))
        return df
    except:
        return pd.DataFrame()


def hybrid_gpt_extraction(file_bytes):
    """Try pdfplumber → if empty then GPT."""
    df = extract_table_pdfplumber(file_bytes)
    if not df.empty:
        return df, "pdfplumber"

    # If no table found → GPT fallback
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        raw_text = "\n".join([page.extract_text() or "" for page in pdf.pages])

    df_gpt = gpt_extract_table(raw_text)
    if not df_gpt.empty:
        return df_gpt, "gpt"

    return pd.DataFrame(), "none"


# ==========================================================
# FINAL CLASSIFICATION RULES
# ==========================================================

def classify_entry(row):
    referencia = str(row.get("Referencia", "")).strip()

    debit = clean_amount(row.get("Debit"))
    credit = clean_amount(row.get("Credit"))

    # RULE 1 → Payment (NO referencia)
    if referencia == "" or referencia.lower() == "none":
        amount = debit if debit is not None else credit
        return {
            "Document": "",
            "Reason": "Payment",
            "Amount": amount
        }

    # RULE 2 → Invoice (Referencia + Debit)
    if debit is not None:
        return {
            "Document": referencia,
            "Reason": "Invoice",
            "Amount": debit
        }

    # RULE 3 → Credit Note (Referencia + Credit)
    if credit is not None:
        return {
            "Document": referencia,
            "Reason": "Credit Note",
            "Amount": credit
        }

    return {
        "Document": referencia,
        "Reason": "Unknown",
        "Amount": None
    }


# ==========================================================
# UI
# ==========================================================

uploaded = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded:
    st.success("📄 PDF uploaded — extracting...")

    file_bytes = uploaded.read()

    df_raw, method = hybrid_gpt_extraction(file_bytes)

    st.write(f"### Extraction Method Used: **{method.upper()}**")

    if df_raw.empty:
        st.error("❌ No table extracted from PDF.")
        st.stop()

    st.write("### Extracted Raw Table")
    st.dataframe(df_raw, use_container_width=True)

    df = df_raw.rename(columns=lambda x: str(x).strip().replace(" ", "_"))

    # Ensure required fields
    for col in ["Referencia", "Debit", "Credit"]:
        if col not in df.columns:
            df[col] = ""

    # Classification
    results = [classify_entry(row) for _, row in df.iterrows()]
    final_df = pd.DataFrame(results)

    st.write("### 🧠 Classified Results (FINAL)")
    st.dataframe(final_df, use_container_width=True)

    # Download
    output = BytesIO()
    final_df.to_excel(output, index=False)

    st.download_button(
        label="⬇️ Download Excel",
        data=output.getvalue(),
        file_name="DataFalcon_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
