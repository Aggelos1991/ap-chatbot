import os, re, json
from io import BytesIO
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openai import OpenAI

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon — GPT-4o-mini DEBE Extractor", layout="wide")
st.title("🦅 DataFalcon — Vendor Statement Extractor (GPT-4o-mini)")

# Load API key
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
MODEL = "gpt-4o-mini"

# ==========================================================
# HELPERS
# ==========================================================
def extract_text_from_pdf(file):
    file_bytes = file.getvalue()
    if not file_bytes:
        raise ValueError("Uploaded file is empty.")
    text = ""
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page in doc:
            text += page.get_text("text") + "\n"
    return text


def normalize_number(value):
    """Normalize all decimal formats (e.g. 1.234,56 → 1234.56 or 1,234.56 → 1234.56)."""
    if not value:
        return ""
    s = str(value).strip()
    s = s.replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s and not "." in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return round(float(s), 2)
    except:
        return ""


def extract_tax_id(text):
    match = re.search(r"\b([A-Z]{1}\d{7}[A-Z0-9]{1}|ES\d{9}|EL\d{9}|[A-Z]{2}\d{8,12})\b", text)
    return match.group(0) if match else "Missing TAX ID"


def extract_with_gpt(raw_text):
    """Send PDF text to GPT-4o-mini and extract DEBE-only lines."""
    prompt = f"""
You are an expert Spanish accountant.

From the following vendor statement, extract all valid invoice or credit note lines.

Each record must include:
- "Alternative Document": invoice or reference (e.g. 6--483, 2434-1, SerieFactura-Precodigo-Num FactCliente)
- "Date": dd/mm/yy or dd/mm/yyyy
- "Reason": "Invoice" or "Credit Note"
- "Document Value": DEBE / IMPORTE / TOTAL (never SALDO or HABER)

Rules:
- Ignore 'Saldo anterior', 'Banco', 'Cobro', 'Efecto', 'Remesa', 'Pago'.
- If the text mentions 'Abono', 'Nota de crédito', or 'NC', mark as Credit Note and make the amount negative.
- Always output a valid JSON array.
- Use '.' for decimals, no commas or €.

Text:
\"\"\"{raw_text[:15000]}\"\"\"
"""

    try:
        response = client.responses.create(model=MODEL, input=prompt)
        content = response.output_text.strip()
    except Exception as e:
        st.error(f"❌ GPT call failed: {e}")
        return []

    try:
        json_match = re.search(r"\[.*\]", content, re.DOTALL)
        json_text = json_match.group(0) if json_match else content
        data = json.loads(json_text)
    except Exception as e:
        st.error(f"⚠️ Could not parse GPT output: {e}")
        st.text_area("🔍 Raw GPT Output", content[:2000], height=200)
        return []

    cleaned = []
    for row in data:
        val = normalize_number(row.get("Document Value") or row.get("DocumentValue"))
        if val == "":
            continue
        reason = row.get("Reason", "").lower()
        if any(k in reason for k in ["abono", "nota de crédito", "credit note", "nc"]):
            val = -abs(val)
            reason = "Credit Note"
        else:
            reason = "Invoice"
        cleaned.append({
            "Alternative Document": row.get("Alternative Document", "").strip(),
            "Date": row.get("Date", "").strip(),
            "Reason": reason,
            "Document Value": val
        })
    return cleaned


def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT APP
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from PDF..."):
        try:
            text = extract_text_from_pdf(uploaded_pdf)
        except Exception as e:
            st.error(f"❌ PDF extraction failed: {e}")
            st.stop()

    approx_tokens = len(text) / 4
    est_cost = (approx_tokens / 1000) * (0.0006 + 0.0024)
    st.info(f"💲 Estimated cost for this extraction: **${est_cost:.4f} USD**")

    if st.button("🤖 Extract Data with GPT-4o-mini"):
        with st.spinner("Analyzing with GPT-4o-mini... please wait..."):
            data = extract_with_gpt(text)

        tax_id = extract_tax_id(text)
        for row in data:
            row["Tax ID"] = tax_id

        if not data:
            st.warning("⚠️ No valid invoice data found. Verify the PDF format or retry.")
        else:
            df = pd.DataFrame(data)
            st.success(f"✅ Extraction complete — {len(df)} invoices found (DEBE-only, Saldo/Haber ignored).")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name="vendor_statement_gpt4omini.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Please upload a vendor statement PDF to begin.")
