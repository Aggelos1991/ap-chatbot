import os, json, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# =============================================
# ENVIRONMENT SETUP
# =============================================
try:
    from dotenv import load_dotenv
    load_dotenv()
except ModuleNotFoundError:
    st.warning("⚠️ 'python-dotenv' not installed — continuing without .env support.")

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No OpenAI API key found. Add it to .env or Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# =============================================
# STREAMLIT CONFIG
# =============================================
st.set_page_config(page_title="🦅 DataFalcon — Vendor Statement Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor (Saldo-Proof + Totale Edition)")

# =============================================
# HELPERS
# =============================================
def extract_text_from_pdf(file):
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())

def normalize_number(value):
    if not value:
        return ""
    s = str(value).strip()
    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^\d{1,3}(,\d{3})*\.\d{2}$", s):
        s = s.replace(",", "")
    elif re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")
    else:
        s = re.sub(r"[^\d.-]", "", s)
    return s

def extract_tax_id(raw_text):
    patterns = [
        r"\b[A-Z]{1}\d{7}[A-Z0-9]{1}\b",
        r"\bES\d{9}\b",
        r"\bEL\d{9}\b",
        r"\b[A-Z]{2}\d{8,12}\b",
    ]
    for pat in patterns:
        match = re.search(pat, raw_text)
        if match:
            return match.group(0)
    return None

# =============================================
# CORE EXTRACTION — SALDO-PROOF
# =============================================
def extract_with_llm(raw_text):
    """
    GPT identifies structure; regex extracts numeric values from DEBE / IMPORTE / VALOR / TOTAL / TOTALE / AMOUNT only.
    SALDO, HABER, BALANCE, payments always ignored.
    """
    # 1️⃣ GPT → detect document numbers & dates
    prompt = f"""
    You are a Spanish accountant AI.
    Read the following vendor statement and identify only document lines (invoices or credit notes).

    For each document include:
    - Alternative Document (Factura / Documento / No / Nº / Num / Número / Nro / Doc)
    - Date (Fecha)
    - Reason ("Invoice" or "Credit Note")

    Ignore:
    - SALDO, BALANCE, ACUMULADO, RESTANTE
    - HABER, CRÉDITO, PAGO, BANCO, REMESA, COBRO, DOMICILIACIÓN
    - Any totals unrelated to documents

    Output JSON with:
    ["Alternative Document", "Date", "Reason"]

    Text:
    \"\"\"{raw_text[:24000]}\"\"\"
    """
    try:
        response = client.responses.create(model=MODEL, input=prompt)
        content = response.output_text.strip()
        json_match = re.search(r"\[.*\]", content, re.DOTALL)
        content = json_match.group(0) if json_match else content
        gpt_rows = json.loads(content)
    except Exception:
        gpt_rows = []

    # 2️⃣ Regex → capture numeric DEBE / TOTAL / IMPORTE / VALOR values, ignoring SALDO & HABER
    pattern = re.compile(
        r"(6[-–]\d{1,4})"                                   # Document number
        r".{0,60}?"                                          # up to 60 chars between
        r"(\d{1,2}/\d{1,2}/\d{2,4})"                        # Date
        r".{0,80}?"                                          # allow some text
        r"(?:DEBE|DEBITO|CARGO|IMPORTE|VALOR|TOTAL|TOTALE|AMOUNT)[\s:=]*([\d.,]{3,10})"  # capture value
        r"(?![\s]*(SALDO|BALANCE|HABER|CR[EÉ]DITO))",       # exclude saldo/haber
        re.IGNORECASE
    )

    regex_data = []
    for m in pattern.finditer(raw_text):
        doc, date, val, _ = m.groups()
        num = normalize_number(val)
        try:
            amount = float(num)
        except:
            continue
        if amount <= 0 or amount > 100000:  # reject ledger codes
            continue
        regex_data.append({
            "Alternative Document": doc.strip(),
            "Date": date.strip(),
            "Document Value": f"{amount:.2f}"
        })

    # 3️⃣ Merge GPT structure (Reason) with regex numeric data
    tax_id = extract_tax_id(raw_text)
    merged = []
    for r in regex_data:
        doc, date, val = r["Alternative Document"], r["Date"], r["Document Value"]
        reason = "Invoice"
        for g in gpt_rows:
            if doc in str(g.get("Alternative Document", "")) or date in str(g.get("Date", "")):
                reason = g.get("Reason", "Invoice")
                break
        merged.append({
            "Alternative Document": doc,
            "Date": date,
            "Reason": reason,
            "Document Value": val,
            "Tax ID": tax_id if tax_id else "Missing TAX ID"
        })

    df = pd.DataFrame(merged).drop_duplicates(subset=["Alternative Document", "Date"])
    return df.to_dict(orient="records")

# =============================================
# EXCEL EXPORT
# =============================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# =============================================
# STREAMLIT UI
# =============================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded_pdf))

    st.text_area("🔍 Extracted Text Preview", text[:2000], height=200)

    if st.button("🤖 Extract Data to Excel"):
        with st.spinner("Analyzing..."):
            data = extract_with_llm(text)

        if data:
            df = pd.DataFrame(data)
            st.success("✅ Extraction complete — Only DEBE / IMPORTE / VALOR / TOTAL lines included. SALDO ignored.")
            st.dataframe(df, use_container_width=True)

            excel_bytes = to_excel_bytes(data)
            st.download_button(
                "⬇️ Download Excel (Vendor Statement)",
                data=excel_bytes,
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No valid document data found. Verify your PDF format.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
