import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — GPT Strict Referencia", layout="wide")
st.title("🦅 DataFalcon Pro — GPT Extractor (STRICT REFERENCIA MODE)")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ Missing OPENAI API key.")
    st.stop()

client = OpenAI(api_key=api_key)

MODEL = "gpt-4o-mini"


# ==========================================================
# PDF extraction
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean:
                        all_lines.append(clean)
            else:
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=260,
                                                first_page=idx, last_page=idx)
                    ocr_text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean:
                            all_lines.append(clean)
                except:
                    pass

    return all_lines


# ==========================================================
# STRICT reference extraction
# ==========================================================
def extract_referencia(line):
    matches = re.findall(r"\b\d{12,18}\b", line)
    return matches[0] if matches else ""


# ==========================================================
# GPT parse → ONLY Concepto, Date, Debit, Credit
# ==========================================================
def gpt_extract(lines):
    BATCH = 40
    output = []

    for i in range(0, len(lines), BATCH):
        batch = lines[i:i + BATCH]
        text_block = "\n".join(batch)

        prompt = f"""
Extract ledger rows. Return ONLY:

- Concepto
- Date
- Debit (DEBE)
- Credit (HABER)

❌ Do NOT extract or guess invoice numbers.
❌ Do NOT infer FP/IR.
❌ Do NOT use description tokens as reference.
❌ Do NOT return saldo.

Strict JSON array, no explanation.

Text:
{text_block}
"""

        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[{"role":"user","content":prompt}],
                temperature=0
            )
            content = resp.choices[0].message.content
        except Exception as e:
            st.error(f"GPT error: {e}")
            return []

        # Extract JSON
        m = re.search(r"\[.*\]", content, re.DOTALL)
        if not m:
            continue

        try:
            data = json.loads(m.group(0))
            output.extend(data)
        except:
            continue

    return output


# ==========================================================
# Merge GPT + Referencia
# ==========================================================
def merge_rows(lines, gpt_rows):
    final = []

    limit = min(len(lines), len(gpt_rows))

    for i in range(limit):
        raw   = lines[i]
        row   = gpt_rows[i]

        ref = extract_referencia(raw)
        concepto = row.get("Concepto", "")
        fecha    = row.get("Date", "")
        debit    = normalize(row.get("Debit", ""))
        credit   = normalize(row.get("Credit", ""))

        # Classification
        if ref and debit:
            reason = "Invoice"
        elif ref and credit:
            reason = "Credit Note"
        elif not ref and credit:
            reason = "Payment"
        else:
            continue

        final.append({
            "Referencia": ref,
            "Concepto": concepto,
            "Date": fecha,
            "Reason": reason,
            "Debit": debit,
            "Credit": credit
        })

    return final


# ==========================================================
# Normalize amounts
# ==========================================================
def normalize(v):
    if not v:
        return ""
    s = str(v).replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        return round(float(s), 2)
    except:
        return ""


# ==========================================================
# EXPORT
# ==========================================================
def to_excel(records):
    df = pd.DataFrame(records)
    buff = BytesIO()
    df.to_excel(buff, index=False)
    buff.seek(0)
    return buff


# ==========================================================
# UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload PDF", type=["pdf"])

if uploaded_pdf:
    lines = extract_raw_lines(uploaded_pdf)

    st.text_area("📄 Raw Preview", "\n".join(lines[:30]), height=260)

    if st.button("🚀 Extract"):
        gpt_rows = gpt_extract(lines)
        final = merge_rows(lines, gpt_rows)

        df = pd.DataFrame(final)
        st.dataframe(df, use_container_width=True, hide_index=True)

        if not df.empty:
            total_debit  = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
            total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
            net = total_debit - total_credit

            c1, c2, c3 = st.columns(3)
            c1.metric("Debit", f"{total_debit:,.2f}")
            c2.metric("Credit", f"{total_credit:,.2f}")
            c3.metric("Net", f"{net:,.2f}")

            st.download_button(
                "⬇️ Download Excel",
                to_excel(final),
                "statement.xlsx"
            )

else:
    st.info("Upload a PDF to begin.")
