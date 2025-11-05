import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract

st.set_page_config(page_title="DataFalcon Pro", layout="wide")
st.title("DataFalcon Pro — Hybrid OCR + Grok")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("GROK_API_KEY") or st.secrets.get("GROK_API_KEY")
if not api_key:
    st.error("Add GROK_API_KEY to Streamlit Secrets")
    st.stop()

# ONLY CHANGE: REAL GROK API
from openai import OpenAI
client = OpenAI(api_key=api_key, base_url="https://api.x.ai/v1")
PRIMARY_MODEL = BACKUP_MODEL = "grok-beta"   # ONLY MODEL THAT WORKS

def extract_text_with_ocr(uploaded_pdf):
    all_lines, ocr_pages = [], []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and len(text.strip()) > 10:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean: all_lines.append(clean)
            else:
                ocr_pages.append(i)
                try:
                    img = convert_from_bytes(pdf_bytes, dpi=300, first_page=i, last_page=i)[0]
                    ocr = pytesseract.image_to_string(img, lang="spa+eng+ell")
                    for line in ocr.split("\n"):
                        clean = " ".join(line.split())
                        if clean: all_lines.append(clean)
                except: pass
    return all_lines, ocr_pages

def normalize_number(v):
    if not v: return ""
    s = re.sub(r"[^\d.,-]", "", str(v).replace(" ", ""))
    s = s.replace(",", ".") if s.count(",") == 1 and s.count(".") <= 1 else s.replace(".", "").replace(",", ".")
    try: return round(float(s), 2)
    except: return ""

def parse_gpt_response(content, n):
    m = re.search(r'\[.*\]', content, re.DOTALL)
    if not m: return []
    try: return json.loads(m.group(0))
    except: return []

def extract_with_gpt(lines):
    BATCH = 50
    records = []
    for i in range(0, len(lines), BATCH):
        block = "\n".join(lines[i:i+BATCH])
        prompt = f"""Extract JSON only:
[
  {{"Alternative Document":"...", "Date":"...", "Reason":"Invoice|Payment|Credit Note", "Debit":0, "Credit":0, "Balance":0}}
]
Text:
{block}"""

        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0,
                    max_tokens=4000
                )
                data = parse_gpt_response(resp.choices[0].message.content, i//BATCH+1)
                if data:
                    for r in data:
                        doc = str(r.get("Alternative Document","")).strip()
                        if not doc: continue
                        records.append({
                            "Alternative Document": doc,
                            "Date": str(r.get("Date","")).strip(),
                            "Reason": ["Invoice","Payment","Credit Note"][
                                0 if "cobro|pago|remesa" not in str(r).lower() else
                                1 if "abono|crédit" not in str(r).lower() else 2
                            ],
                            "Debit": normalize_number(r.get("Debit")),
                            "Credit": normalize_number(r.get("Credit")),
                            "Balance": normalize_number(r.get("Balance"))
                        })
                    break
            except Exception as e:
                st.warning(f"Grok error: {e}")
    return records

def to_excel_bytes(recs):
    buf = BytesIO()
    pd.DataFrame(recs).to_excel(buf, index=False)
    buf.seek(0)
    return buf

# UI
pdf_file = st.file_uploader("Upload PDF", type="pdf")
if pdf_file:
    with st.spinner("Reading PDF..."):
        lines, _ = extract_text_with_ocr(pdf_file)
    st.success(f"{len(lines)} lines")
    if st.button("Run Grok Extraction", type="primary"):
        with st.spinner("Grok is flying..."):
            data = extract_with_gpt(lines)
        if data:
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
            st.download_button("Download Excel", 
                data=to_excel_bytes(data),
                file_name=f"extract_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("No data extracted")
else:
    st.info("Upload a PDF to start")
