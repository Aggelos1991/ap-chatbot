import os, json, re
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# =============================================
# ENVIRONMENT & MODEL SETUP
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
st.set_page_config(page_title="📄 DataFalcon — Vendor Statement Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor (Saldo-Proof Version)")

# =============================================
# HELPERS
# =============================================
def extract_text_from_pdf(file):
    """Extract text from PDF."""
    text = ""
    with fitz.open(stream=file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text


def clean_text(text):
    return " ".join(text.replace("\xa0", " ").replace("€", " EUR").split())


def normalize_number(value):
    """Normalize Spanish/EU formatted numbers into float-compatible strings."""
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
    """Detect Spanish CIF/NIF or European VAT patterns."""
    patterns = [
        r"\b[A-Z]{1}\d{7}[A-Z0-9]{1}\b",  # B12345678
        r"\bES\d{9}\b",  # ES123456789
        r"\bEL\d{9}\b",  # EL123456789
        r"\b[A-Z]{2}\d{8,12}\b",  # EU VAT generic
    ]
    for pat in patterns:
        match = re.search(pat, raw_text)
        if match:
            return match.group(0)
    return None

# =============================================
# CORE EXTRACTION LOGIC
# =============================================
def extract_with_llm(raw_text):
    """
    Hybrid extraction: GPT for semantic context + regex fallback for missed DEBE lines.
    """
    # ---------- GPT extraction ----------
    prompt = f"""
    You are an expert accountant AI.

    The following text is a Spanish vendor statement with columns:
    ASIENTO | FECHA | DOC | DEBE | HABER | SALDO

    Each line may look like:
    6--483 24/01/25 Fra. emitida nº 70000000 CREADA POR IMPORTADOR 322,27 386,16 708,43

    Instructions:
    - Extract all document lines (Factura / Credit Note) using DEBE as Document Value.
    - The DEBE value is always the number appearing **before HABER and SALDO**.
    - Ignore "Saldo", "Balance", "Acumulado", "Restante" values.
    - Ignore payments ("Pago", "Transferencia", "Banco", "Cobro Efecto", "Remesa", "Domiciliación").
    - Keep all Factura lines, even if repeated pattern.
    - Output valid JSON with: ["Alternative Document", "Date", "Reason", "Document Value", "Tax ID"]

    Example:
    [
      {{
        "Alternative Document": "6--483",
        "Date": "24/01/25",
        "Reason": "Invoice",
        "Document Value": "322.27",
        "Tax ID": "B12345678"
      }},
      {{
        "Alternative Document": "6--2322",
        "Date": "12/03/25",
        "Reason": "Invoice",
        "Document Value": "132.57",
        "Tax ID": "B12345678"
      }}
    ]

    Text:
    \"\"\"{raw_text[:24000]}\"\"\"
    """

    response = client.responses.create(model=MODEL, input=prompt)
    content = response.output_text.strip()

    try:
        json_match = re.search(r'\[.*\]', content, re.DOTALL)
        if json_match:
            content = json_match.group(0)
        data = json.loads(content)
    except Exception as e:
        st.error(f"⚠️ Could not parse GPT output: {e}")
        data = []

    # ---------- Post-cleanup ----------
    tax_id = extract_tax_id(raw_text)
    filtered = []

    for row in data:
        val = normalize_number(row.get("Document Value", ""))
        if not val:
            continue

        try:
            amount = float(val)
        except:
            continue

        # ignore large / invalid / SALDO / account codes
        if amount <= 0 or amount > 100000:
            continue

        reason = str(row.get("Reason", "")).lower()
        if any(w in reason for w in ["saldo", "balance", "acumulad", "restante", "pago", "transferencia", "banco", "remesa", "cobro"]):
            continue

        if any(w in reason for w in ["abono", "nota de crédito", "nota credito", "devolucion"]):
            amount = -abs(amount)
            row["Reason"] = "Credit Note"
        else:
            row["Reason"] = "Invoice"

        row["Document Value"] = f"{amount:.2f}"
        row["Tax ID"] = tax_id if tax_id else "Missing TAX ID"
        filtered.append(row)

    # ---------- Regex fallback ----------
    regex_matches = re.findall(r'(\d{1,2}/\d{1,2}/\d{2,4}).*?(6[-–]\d{1,4}).*?\s(\d{1,3}[.,]\d{2})\s', raw_text)
    for date, doc, val in regex_matches:
        val = normalize_number(val)
        if not val:
            continue
        if not any(doc == r.get("Alternative Document") for r in filtered):
            filtered.append({
                "Alternative Document": doc,
                "Date": date,
                "Reason": "Invoice",
                "Document Value": val,
                "Tax ID": tax_id if tax_id else "Missing TAX ID"
            })

    return filtered
# =============================================
# EXCEL EXPORT
# =============================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

# =============================================
# STREAMLIT UI
# =============================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        text = clean_text(extract_text_from_pdf(uploaded_pdf))

    st.text_area("🔍 Extracted Text Preview", text[:2000], height=200)

    if st.button("🤖 Extract Data to Excel"):
        with st.spinner("Analyzing with GPT... please wait..."):
            data = extract_with_llm(text)

        if data:
            df = pd.DataFrame(data)
            st.success("✅ Extraction complete — only DEBE (document values) captured, SALDO ignored!")
            st.dataframe(df, use_container_width=True)

            excel_bytes = to_excel_bytes(data)
            st.download_button(
                "⬇️ Download Excel (Vendor Statement)",
                data=excel_bytes,
                file_name="vendor_statement_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No valid document data found. Check the PDF layout or content.")
else:
    st.info("Upload a Spanish vendor statement PDF to begin.")
