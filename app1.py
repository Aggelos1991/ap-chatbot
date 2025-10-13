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
st.title("🦅 DataFalcon — Vendor Statement Extractor (Saldo-Proof + Debug Edition)")

# =============================================
# HELPERS
# =============================================
def extract_text_from_pdf(file):
    """Safely extract text from uploaded PDF."""
    file_bytes = file.getvalue()
    if not file_bytes:
        raise ValueError("Uploaded file is empty or unreadable.")

    text = ""
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text


def clean_text(text):
    """Clean extracted text."""
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())


def normalize_number(value):
    """Normalize European/US number formats."""
    if not value:
        return ""
    s = str(value).strip()
    if re.match(r"^\d{1,3}(\.\d{3})*,\d{2}$", s):  # 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif re.match(r"^\d{1,3}(,\d{3})*\.\d{2}$", s):  # 1,234.56
        s = s.replace(",", "")
    elif re.match(r"^\d+,\d{2}$", s):  # 150,00
        s = s.replace(",", ".")
    else:
        s = re.sub(r"[^\d.-]", "", s)
    return s


def extract_tax_id(raw_text):
    """Detect CIF/NIF/VAT from text."""
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
# CORE EXTRACTION (Saldo-Proof + Totale support)
# =============================================
def extract_with_llm(raw_text):
    """
    Hybrid extraction:
    ✅ Extracts DEBE / IMPORTE / VALOR / TOTAL / TOTALE / AMOUNT
    ❌ Ignores SALDO, HABER, BALANCE, payments.
    Works even if no column headers exist.
    """

    # ---------- 1️⃣ GPT extracts document/date/reason ----------
    prompt = f"""
    You are a Spanish accountant AI.
    Identify all invoice or credit note lines from this vendor statement text.

    Each record must contain:
    - Alternative Document (Factura, Documento, No, Nº, Num, Número, Nro, Doc)
    - Date (Fecha)
    - Reason ("Invoice" or "Credit Note")

    Ignore any line containing:
    SALDO, BALANCE, ACUMULADO, RESTANTE,
    HABER, CRÉDITO, PAGO, BANCO, REMESA, COBRO, DOMICILIACIÓN.

    Return a valid JSON list of objects with those three fields.
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

    # ---------- 2️⃣ Regex extract numeric document values ----------
    # Captures DEBE, IMPORTE, VALOR, TOTAL, TOTALE, AMOUNT
    # Ignores SALDO, HABER, BALANCE, PAGO, BANCO, COBRO, REMESA.
    pattern = re.compile(
        r"(?P<doc>(?:\b\d{1,3}[-–]\d{1,5}\b|\b6[-–]\d{1,5}\b)).{0,60}?"
        r"(?P<date>\d{1,2}/\d{1,2}/\d{2,4}).{0,80}?"
        r"(?:(?:DEBE|IMPORTE|VALOR|TOTAL|TOTALE|AMOUNT)[\s:=]*)?"
        r"(?P<amount>[\d.,]{3,10})"
        r"(?![\s]*(SALDO|BALANCE|HABER|CR[EÉ]DITO|PAGO|BANCO|COBRO|REMESA))",
        re.IGNORECASE,
    )

    regex_data = []
    for m in pattern.finditer(raw_text):
        doc, date, val = m.group("doc"), m.group("date"), m.group("amount")
        val = normalize_number(val)
        try:
            amount = float(val)
        except:
            continue
        if amount <= 0 or amount > 100000:
            continue
        regex_data.append(
            {
                "Alternative Document": doc.strip(),
                "Date": date.strip(),
                "Document Value": f"{amount:.2f}",
            }
        )

    # ---------- 3️⃣ Merge GPT structure ----------
    tax_id = extract_tax_id(raw_text)
    merged = []
    for r in regex_data:
        doc, date, val = r["Alternative Document"], r["Date"], r["Document Value"]
        reason = "Invoice"
        for g in gpt_rows:
            if doc in str(g.get("Alternative Document", "")) or date in str(
                g.get("Date", "")
            ):
                reason = g.get("Reason", "Invoice")
                break
        merged.append(
            {
                "Alternative Document": doc,
                "Date": date,
                "Reason": reason,
                "Document Value": val,
                "Tax ID": tax_id if tax_id else "Missing TAX ID",
            }
        )

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
        try:
            text = clean_text(extract_text_from_pdf(uploaded_pdf))
        except Exception as e:
            st.error(f"❌ Failed to read PDF: {e}")
            st.stop()

    # 🧩 DEBUG SECTION (so we can calibrate)
    st.subheader("🧩 Debug: What the PDF text actually looks like")
    st.text_area("Raw Extracted Text (first 4000 chars)", text[:4000], height=300)
    st.download_button("⬇️ Download full extracted text", text.encode("utf-8"), "raw_text.txt")

    # ---------- Extraction ----------
    if st.button("🤖 Extract Data to Excel"):
        with st.spinner("Analyzing with GPT + Regex..."):
            data = extract_with_llm(text)

        if data:
            df = pd.DataFrame(data)
            st.success(
                "✅ Extraction complete — Only DEBE / IMPORTE / VALOR / TOTAL / TOTALE / AMOUNT used. SALDO ignored."
            )
            st.dataframe(df, use_container_width=True)

            excel_bytes = to_excel_bytes(data)
            st.download_button(
                "⬇️ Download Excel (Vendor Statement)",
                data=excel_bytes,
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning(
                "⚠️ No valid document data found. Copy one or two lines from the debug text above and share them for calibration."
            )
else:
    st.info("Please upload a vendor statement PDF to begin.")
