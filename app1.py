import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro")

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

PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"

# ==========================================================
# PDF LINE CLEANER
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    ocr_pages = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()

            if text:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean and "saldo" not in clean.lower():
                        all_lines.append(clean)
            else:
                ocr_pages.append(idx)
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=260, first_page=idx, last_page=idx)
                    ocr_text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean and "saldo" not in clean.lower():
                            all_lines.append(clean)
                except Exception as e:
                    st.warning(f"OCR skipped page {idx}: {e}")

    return all_lines


# ==========================================================
# GPT PARSER
# ==========================================================
def parse_gpt_response(content):
    m = re.search(r"\[.*\]", content, re.DOTALL)
    if not m:
        return []
    try:
        return json.loads(m.group(0))
    except:
        return []


# ==========================================================
# GPT EXTRACTOR – NOW RETURNS ONLY:
# Concepto, Date, Debit, Credit
# WE ADD Referencia OURSELVES
# ==========================================================
def extract_with_gpt(lines):

    BATCH = 65
    output = []

    for i in range(0, len(lines), BATCH):
        batch = lines[i:i+BATCH]
        text_block = "\n".join(batch)

        prompt = f"""
Extract accounting entries.

Return ONLY:
- Concepto (description text)
- Date (any date found)
- Debit (DEBE)
- Credit (HABER)

IMPORTANT:
❌ DO NOT extract or guess invoice/document numbers.
❌ DO NOT return references like FP, IR, VARIOS, CA000194.
❌ DO NOT create invoice-like codes.

ONLY return the 4 required fields.

Strict JSON array, no explanation.

Text:
{text_block}
"""

        parsed = []

        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                content = resp.choices[0].message.content.strip()
                parsed = parse_gpt_response(content)
                if parsed:
                    break
            except Exception as e:
                st.warning(f"GPT error on model {model}: {e}")

        if not parsed:
            continue

        # CLEAN RECORDS
        for row in parsed:
            concepto = str(row.get("Concepto", "")).strip()
            date = str(row.get("Date", "")).strip()
            debit = normalize(row.get("Debit", ""))
            credit = normalize(row.get("Credit", ""))

            if not debit and not credit:
                continue

            output.append({
                "Concepto": concepto,
                "Date": date,
                "Debit": debit,
                "Credit": credit
            })

    return output


# ==========================================================
# NORMALIZE NUMBERS
# ==========================================================
def normalize(v):
    if not v:
        return ""
    s = str(v).replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return round(float(s), 2)
    except:
        return ""


# ==========================================================
# REFERENCIA EXTRACTOR (OUR RULE)
# ==========================================================
def extract_referencia_from_line(line):
    """
    Extract ONLY the official Referencia.
    Pattern is always a long numeric chain:

    Example: 230101183005951
             230126151000009
             230101183005531
    """
    matches = re.findall(r"\b\d{12,18}\b", line)
    if matches:
        return matches[0]
    return ""


# ==========================================================
# MERGE GPT OUTPUT + REFERENCIA
# ==========================================================
def merge_with_referencia(lines, gpt_rows):
    final = []

    for row in gpt_rows:

        concepto = row["Concepto"]

        # FIND the original PDF line that matches this concepto (best-effort)
        ref = ""
        for line in lines:
            if concepto[:20] in line:
                ref = extract_referencia_from_line(line)
                break

        debit = row["Debit"]
        credit = row["Credit"]

        # CLASSIFY
        if debit:
            reason = "Invoice"
        elif credit:
            # Payment or credit note?
            if re.search(r"pago|cobro|transfer", concepto, re.IGNORECASE):
                reason = "Payment"
            else:
                reason = "Credit Note"
        else:
            continue

        final.append({
            "Referencia": ref,
            "Concepto": concepto,
            "Date": row["Date"],
            "Reason": reason,
            "Debit": debit,
            "Credit": credit
        })

    return final


# ==========================================================
# EXPORT EXCEL
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buff = BytesIO()
    df.to_excel(buff, index=False)
    buff.seek(0)
    return buff


# ==========================================================
# STREAMLIT UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    lines = extract_raw_lines(uploaded_pdf)
    st.text_area("Preview", "\n".join(lines[:25]), height=250)

    if st.button("🚀 Extract"):
        gpt_rows = extract_with_gpt(lines)
        final = merge_with_referencia(lines, gpt_rows)

        df = pd.DataFrame(final)
        st.dataframe(df, hide_index=True, use_container_width=True)

        if not df.empty:
            total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
            total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()

            c1, c2, c3 = st.columns(3)
            c1.metric("Debit", f"{total_debit:,.2f}")
            c2.metric("Credit", f"{total_credit:,.2f}")
            c3.metric("Net", f"{total_debit-total_credit:,.2f}")

            st.download_button(
                "⬇️ Download Excel",
                to_excel_bytes(final),
                "statement.xlsx"
            )
else:
    st.info("Upload a PDF to begin.")
