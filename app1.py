import os, re, json, time, concurrent.futures
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
st.set_page_config(page_title="DataFalcon Pro — FINAL FIX", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL FIXED VERSION")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("Missing OpenAI API Key.")
    st.stop()

client = OpenAI(api_key=api_key)

PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"


# ==========================================================
# HELPERS
# ==========================================================
def normalize(v):
    if not v:
        return ""
    s = str(v).replace(" ", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return round(float(s), 2)
    except:
        return ""


# ==========================================================
# PDF TEXT EXTRACTION
# ==========================================================
def extract_raw_lines(file):
    out = []
    pdf_bytes = file.read()
    file.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for ln in text.split("\n"):
                    ln = ln.strip()
                    if ln and "saldo" not in ln.lower():
                        out.append(ln)
            else:
                try:
                    img = convert_from_bytes(pdf_bytes, dpi=200, first_page=page.page_number, last_page=page.page_number)[0]
                    t = pytesseract.image_to_string(img, lang="spa+eng+ell")
                    for ln in t.split("\n"):
                        ln = ln.strip()
                        if ln and "saldo" not in ln.lower():
                            out.append(ln)
                except:
                    pass

    return out


# ==========================================================
# 1) CLEAN LEDGER → KEEP ONLY LINES STARTING WITH DATE
# ==========================================================
def keep_only_date_lines(lines):
    pattern = r"^\d{2}/\d{2}/\d{4}"
    return [ln for ln in lines if re.match(pattern, ln)]


# ==========================================================
# 2) MERGE MULTI-LINE RECORDS INTO ONE LINE
# ==========================================================
def merge_multiline_rows(lines):
    merged = []
    buffer = ""
    pattern = r"^\d{2}/\d{2}/\d{4}"

    for ln in lines:
        if re.match(pattern, ln):
            if buffer.strip():
                merged.append(buffer.strip())
            buffer = ln
        else:
            buffer += " " + ln.strip()

    if buffer.strip():
        merged.append(buffer.strip())

    return merged


# ==========================================================
# 3) SPLIT MULTIPLE RECORDS INSIDE ONE PHYSICAL LINE
# ==========================================================
def split_multiple_in_one_line(lines):
    final = []
    pattern = r"(?=\d{2}/\d{2}/\d{4})"

    for ln in lines:
        parts = re.split(pattern, ln)
        for p in parts:
            p = p.strip()
            if re.match(r"^\d{2}/\d{2}/\d{4}", p):
                final.append(p)

    return final


# ==========================================================
# GPT TIMEOUT WRAPPER
# ==========================================================
def gpt_timeout(model, prompt, timeout=10):
    def job():
        return client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}]
        )

    with concurrent.futures.ThreadPoolExecutor() as ex:
        f = ex.submit(job)
        try:
            return f.result(timeout=timeout)
        except:
            return None


# ==========================================================
# PARSE GPT JSON
# ==========================================================
def parse_json(x):
    m = re.search(r"\[.*\]", x, re.DOTALL)
    if not m:
        return []
    try:
        return json.loads(m.group(0))
    except:
        return []


# ==========================================================
# GPT LEDGER EXTRACTION
# ==========================================================
def extract_with_gpt(lines):
    out = []
    B = 20

    for i in range(0, len(lines), B):
        block = "\n".join(lines[i:i+B])

        prompt = f"""
Extract ledger rows.

RULES:
- Document number = ONLY 'Referencia'.
- If Referencia empty → Payment.
- If Asiento = VEN AND Credit > 0 → Credit Note.
- Else → Invoice.

FORMAT:
Fecha | Asiento | Documento | Libro | Descripción | Referencia | Debe | Haber

Return ONLY JSON array:
[
  {{
    "Fecha": "",
    "Asiento": "",
    "Referencia": "",
    "Concepto": "",
    "Debit": "",
    "Credit": ""
  }}
]

Text:
{block}
"""

        response = gpt_timeout(PRIMARY_MODEL, prompt, timeout=10)
        if response is None:
            response = gpt_timeout(BACKUP_MODEL, prompt, timeout=10)

        if response is None:
            st.warning(f"⚠️ GPT timeout on batch {i//B+1}, skipped.")
            continue

        rows = parse_json(response.choices[0].message.content)
        if not rows:
            continue

        for r in rows:
            ref = r.get("Referencia", "").strip()
            asiento = r.get("Asiento", "").strip().upper()

            debit = normalize(r.get("Debit"))
            credit = normalize(r.get("Credit"))

            if ref == "":
                reason = "Payment"
            elif asiento == "VEN" and credit not in ("", 0) and float(credit) > 0:
                reason = "Credit Note"
            else:
                reason = "Invoice"

            out.append({
                "Document": ref,
                "Date": r.get("Fecha", ""),
                "Asiento": asiento,
                "Concepto": r.get("Concepto", ""),
                "Reason": reason,
                "Debit": debit,
                "Credit": credit
            })

    return out


# ==========================================================
# EXCEL EXPORT
# ==========================================================
def to_excel(df):
    b = BytesIO()
    df.to_excel(b, index=False)
    b.seek(0)
    return b


# ==========================================================
# UI
# ==========================================================
uploaded = st.file_uploader("Upload Ledger PDF", type=["pdf"])

if uploaded:
    # 1) Raw extraction
    raw = extract_raw_lines(uploaded)

    # 2) Keep only ledger lines starting with date
    lines = keep_only_date_lines(raw)

    # 3) Merge broken multi-line records
    lines = merge_multiline_rows(lines)

    # 4) Split multiple entries hidden inside one line
    lines = split_multiple_in_one_line(lines)

    st.success(f"Detected {len(lines)} ledger rows.")
    st.text_area("Preview", "\n".join(lines[:40]), height=250)

    if st.button("Run DataFalcon Extraction"):
        with st.spinner("Processing with GPT…"):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success("Extraction complete!")
            st.dataframe(df, hide_index=True, use_container_width=True)

            st.download_button(
                "Download Excel",
                data=to_excel(df),
                file_name="datafalcon_output.xlsx"
            )
        else:
            st.error("No rows extracted.")

else:
    st.info("Upload a PDF to begin.")
