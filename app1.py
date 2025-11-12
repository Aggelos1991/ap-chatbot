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
st.set_page_config(page_title="🦅 DataFalcon Pro — OCR + Regex Hybrid", layout="wide")
st.title("🦅 DataFalcon Pro — Final OCR + Regex Accounting Edition")
try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass
api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No OpenAI API key found.")
    st.stop()
client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"
# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(v):
    if pd.isna(v) or v is None or (isinstance(v, str) and str(v).strip() == ""):
        return ""
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    s = str(v).replace("€", "").replace(" ", "").replace("$", "").replace("£", "")
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
    except ValueError:
        return ""
# ==========================================================
# PDF + OCR
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    ocr_pages = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text or len(text.strip()) < 10:
                ocr_pages.append(i)
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=300, first_page=i, last_page=i)
                    text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                except Exception as e:
                    st.warning(f"OCR failed on page {i}: {e}")
                    continue
            for line in text.split("\n"):
                clean = " ".join(line.split())
                if not clean:
                    continue
                if re.search(r"\b(saldo|balance|total\s*saldo|anterior|final|total)\b", clean, re.IGNORECASE):
                    continue
                all_lines.append(clean)
    if ocr_pages:
        st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")
    return all_lines
# ==========================================================
# GPT → Structure (doc/date/desc/debe/haber)
# ==========================================================
def gpt_structure(lines):
    all_records = []
    BATCH = 40  # Smaller batch to avoid token limits
    for i in range(0, len(lines), BATCH):
        batch = lines[i:i+BATCH]
        text_block = "\n".join(batch)
        prompt = f"""
You are an expert accounting statement parser. The text consists of lines from a vendor statement table with columns: Fecha, Dia, Ord., Concepto, Docum., Debe, Haber, Saldo.
Identify only transaction lines (ignore headers, footers, summaries, saldo anterior or final).
For each transaction, extract:
- Alternative Document: the Docum. field (e.g., NF A25021)
- Date: the Fecha field (e.g., 31/01/25)
- Description: the Concepto field (e.g., N.F. A250213 or Cobro factura A250269 Rec)
- Debe: the Debe amount exactly as string (e.g., "907,98") or empty string "" if no amount in Debe
- Haber: the Haber amount exactly as string (e.g., "542,90") or empty string "" if no amount in Haber
IMPORTANT: Do NOT extract or include the Saldo or balance column in Debe or Haber. Ignore Saldo completely. Set Debe or Haber to "" if the column is empty.
Output only a valid JSON array of objects, nothing else.
Example:
Text:
31/01/25 1 245 N.F. A250213 NF A25021 907,98  6.355,74
26/02/25 1 801 Cobro factura A250269 Rec NF A25069  542,90 3.719,83
Output:
[
  {{
    "Alternative Document": "NF A25021",
    "Date": "31/01/25",
    "Description": "N.F. A250213",
    "Debe": "907,98",
    "Haber": ""
  }},
  {{
    "Alternative Document": "NF A25069",
    "Date": "26/02/25",
    "Description": "Cobro factura A250269 Rec",
    "Debe": "",
    "Haber": "542,90"
  }}
]
Text:
{text_block}
"""
        try:
            r = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            content = r.choices[0].message.content.strip()
            if content.startswith('[') and content.endswith(']'):
                data = json.loads(content)
            else:
                match = re.search(r'\[[\s\S]*\]', content)
                if match:
                    data = json.loads(match.group(0))
                else:
                    raise ValueError("No JSON array found")
            all_records.extend(data)
        except Exception as e:
            st.warning(f"GPT batch {i//BATCH+1} failed: {e}")
    return all_records
# ==========================================================
# CLASSIFY BY POLARITY
# ==========================================================
def classify(records):
    parsed = []
    for r in records:
        doc = str(r.get("Alternative Document", "")).strip()
        date = str(r.get("Date", "")).strip()
        desc = str(r.get("Description", "")).strip()
        debe_raw = str(r.get("Debe", "")).strip()
        haber_raw = str(r.get("Haber", "")).strip()
        debe = normalize_number(debe_raw)
        haber = normalize_number(haber_raw)
        if debe == "" and haber == "":
            continue
        reason = ""
        amount = 0.0
        if debe != "" and haber == "":
            sign = 1 if debe > 0 else -1
            reason = "Invoice" if sign > 0 else "Credit Note"
            amount = abs(debe)
        elif haber != "" and debe == "":
            sign = 1 if haber > 0 else -1
            reason = "Payment" if sign > 0 else "Reversal"
            amount = abs(haber)
        elif debe != "" and haber != "":
            reason = "Payment"  # or handle mixed
            amount = abs(haber) if haber != 0 else abs(debe)
        parsed.append({
            "Alternative Document": doc,
            "Date": date,
            "Description": desc,
            "Debe": debe,
            "Haber": haber,
            "Reason": reason,
            "Amount": amount
        })
    return pd.DataFrame(parsed)
# ==========================================================
# STREAMLIT APP
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])
if uploaded_pdf:
    with st.spinner("🧩 Extracting text …"):
        lines = extract_raw_lines(uploaded_pdf)
    st.success(f"✅ {len(lines)} lines extracted.")
    st.text_area("Preview of text (30 lines):", "\n".join(lines[:30]), height=250)
    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Parsing structure with GPT …"):
            base = gpt_structure(lines)
            df = classify(base)
        if len(df) == 0:
            st.warning("⚠️ No records found.")
            st.stop()
        st.success(f"✅ Extraction complete — {len(df)} records found.")
        st.dataframe(df, use_container_width=True, hide_index=True)
        totals = df.groupby("Reason")["Amount"].sum().round(2).reset_index()
        st.markdown("### 💰 Summary by Type")
        st.dataframe(totals, hide_index=True)
        buf = BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)
        st.download_button(
            "⬇️ Download Excel",
            data=buf,
            file_name=f"DataFalcon_Hybrid_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload a vendor statement PDF to begin.")
