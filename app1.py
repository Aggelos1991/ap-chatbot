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
st.set_page_config(page_title="🦅 DataFalcon Pro — FINAL VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL VERSION")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ Missing OpenAI API Key.")
    st.stop()

client = OpenAI(api_key=api_key)
PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"


# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(v):
    if not v:
        return ""
    s = str(v).replace(" ", "").strip()
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


# ==========================================================
# PDF + OCR EXTRACTION
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean and not re.search(r"saldo", clean, re.IGNORECASE):
                        all_lines.append(clean)
            else:
                try:
                    img = convert_from_bytes(pdf_bytes, dpi=240, first_page=i, last_page=i)[0]
                    ocr_text = pytesseract.image_to_string(img, lang="spa+eng+ell")
                    for ln in ocr_text.split("\n"):
                        clean = " ".join(ln.split())
                        if clean and not re.search(r"saldo", clean, re.IGNORECASE):
                            all_lines.append(clean)
                except:
                    pass

    return all_lines


# ==========================================================
# **THE CRITICAL FIX**
# SPLIT MULTIPLE RECORDS IN THE SAME PHYSICAL LINE
# ==========================================================
def preprocess_ledger_lines(raw_lines):
    final = []
    pattern = r"\d{2}/\d{2}/\d{4}"

    for ln in raw_lines:
        parts = re.split(f"(?=\\b{pattern}\\b)", ln)
        for p in parts:
            p = p.strip()
            if re.match(pattern, p):
                final.append(p)

    return final


# ==========================================================
# PARSE GPT JSON
# ==========================================================
def parse_gpt_response(content):
    m = re.search(r'\[.*\]', content, re.DOTALL)
    if not m:
        return []
    try:
        return json.loads(m.group(0))
    except:
        return []


# ==========================================================
# GPT TIMEOUT WRAPPER
# ==========================================================
def gpt_call_with_timeout(model, prompt, timeout=12):
    def task():
        return client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}]
        )

    with concurrent.futures.ThreadPoolExecutor() as executor:
        future = executor.submit(task)
        try:
            return future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            return None


# ==========================================================
# MAIN GPT EXTRACTOR
# ==========================================================
def extract_with_gpt(lines):
    all_records = []
    BATCH = 20  # fast + stable

    for i in range(0, len(lines), BATCH):
        block = "\n".join(lines[i:i+BATCH])

        prompt = f"""
Extract ledger entries.

RULES:
- Document number = ONLY 'Referencia'.
- If Referencia empty → Payment.
- If Asiento = VEN AND Credit > 0 → Credit Note.
- Otherwise → Invoice.
- Do NOT invent numbers.
- Do NOT extract numbers from Concepto.

FORMAT:
Fecha | Asiento | Documento | Libro | Descripción | Referencia | F. valor | Debe | Haber

Return JSON ONLY:
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

TEXT TO ANALYZE:
{block}
"""

        response = gpt_call_with_timeout(PRIMARY_MODEL, prompt, timeout=12)
        if response is None:
            response = gpt_call_with_timeout(BACKUP_MODEL, prompt, timeout=12)

        if response is None:
            st.warning(f"⚠️ GPT timeout on batch {i//BATCH+1}, skipping.")
            continue

        content = response.choices[0].message.content
        data = parse_gpt_response(content)
        if not data:
            continue

        for r in data:
            ref = str(r.get("Referencia", "")).strip()
            asiento = str(r.get("Asiento", "")).strip().upper()

            debit = normalize_number(r.get("Debit", ""))
            credit = normalize_number(r.get("Credit", ""))

            if ref == "":
                reason = "Payment"
            elif asiento == "VEN" and credit not in ("", 0) and float(credit) > 0:
                reason = "Credit Note"
            else:
                reason = "Invoice"

            all_records.append({
                "Document": ref,
                "Date": r.get("Fecha", ""),
                "Asiento": asiento,
                "Concepto": r.get("Concepto", ""),
                "Reason": reason,
                "Debit": debit,
                "Credit": credit
            })

    return all_records


# ==========================================================
# EXPORT
# ==========================================================
def to_excel(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ==========================================================
# UI
# ==========================================================
uploaded = st.file_uploader("📂 Upload Vendor Ledger PDF", type=["pdf"])

if uploaded:
    raw = extract_raw_lines(uploaded)
    lines = preprocess_ledger_lines(raw)

    if not lines:
        st.error("No ledger rows detected.")
        st.stop()

    st.success(f"Detected {len(lines)} real ledger rows.")
    st.text_area("Preview", "\n".join(lines[:40]), height=250)

    if st.button("🚀 Run DataFalcon Extraction", type="primary"):
        with st.spinner("Processing with GPT…"):
            records = extract_with_gpt(lines)

        if records:
            df = pd.DataFrame(records)
            st.success(f"Extraction complete! {len(df)} rows.")
            st.dataframe(df, hide_index=True, use_container_width=True)

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel(df),
                file_name="datafalcon_output.xlsx"
            )
        else:
            st.error("No records extracted.")
else:
    st.info("Upload a PDF to begin.")
