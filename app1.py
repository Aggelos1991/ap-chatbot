import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

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
MODEL = "gpt-4o-mini"

# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(value):
    """Normalize decimals like 1.234,56 → 1234.56"""
    if not value:
        return ""
    s = str(value).strip().replace(" ", "")
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

def extract_raw_lines(uploaded_pdf):
    """Extract all text lines from every page of the PDF."""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", line):
                    all_lines.append(" ".join(line.split()))
    return all_lines

# ==========================================================
# GPT EXTRACTOR — detect DEBE & HABER columns explicitly
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to detect Debit (DEBE) and Credit (HABER) from vendor statements."""
    BATCH_SIZE = 150
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are an expert accountant fluent in Spanish and Greek.

You are reading extracted lines from a vendor statement.
Each line may include columns labeled as:
- DEBE → Debit (Invoice)
- HABER → Credit (Payment)
- SALDO → Running Balance
- CONCEPTO → Description such as "Fra. emitida", "Cobro Efecto", etc.

Your task:
For each valid transaction line, output:
- "Alternative Document": document number (under Nº, Num, Documento, Factura, etc.)
- "Date": date if visible (dd/mm/yy or dd/mm/yyyy)
- "Reason": classify as "Invoice", "Payment", or "Credit Note"
- "Debit": numeric value under DEBE column (if exists)
- "Credit": numeric value under HABER column (if exists)

Rules:
1. If DEBE > 0 → Reason = "Invoice"
2. If HABER > 0 → Reason = "Payment"
3. If the line includes "Abono", "Nota de Credito", "NC", "πιστω", "Ακυρωτικό" → Reason = "Credit Note" and place value under Credit.
4. Ignore summary lines: "Saldo", "Apertura", "Total General", "IVA", "Base", "Impuestos".
5. Exclude any line where the document number contains “concil” (not case-sensitive).
6. Ensure output is valid JSON array.

Lines:
\"\"\"{text_block}\"\"\"
"""

        try:
            response = client.responses.create(model=MODEL, input=prompt)
            content = response.output_text.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            if not json_match:
                continue
            data = json.loads(json_match.group(0))
        except Exception as e:
            st.warning(f"⚠️ GPT failed on batch {i//BATCH_SIZE + 1}: {e}")
            continue

        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            # 🚫 exclude concil. or conciliación or reconcil etc.
            if re.search(r"concil", alt_doc, re.IGNORECASE):
                continue

            debit_val = normalize_number(row.get("Debit"))
            credit_val = normalize_number(row.get("Credit"))

            reason_text = row.get("Reason", "").lower()
            concept = alt_doc.lower()

            # Move Cobro/Efecto to Credit if missing
            if "cobro" in concept or "efecto" in concept:
                credit_val = credit_val or debit_val
                debit_val = ""

            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": row.get("Reason", "").strip(),
                "Debit": debit_val,
                "Credit": credit_val
            })

    return all_records

# ==========================================================
# EXPORT
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from all pages..."):
        lines = extract_raw_lines(uploaded_pdf)

    if not lines:
        st.warning("⚠️ No readable text lines found. Check if the PDF is scanned.")
    else:
        st.text_area("📄 Preview (first 25 lines):", "\n".join(lines[:25]), height=250)

        if st.button("🤖 Run Hybrid Extraction"):
            with st.spinner("Analyzing data with GPT-4o-mini..."):
                data = extract_with_gpt(lines)

            if not data:
                st.warning("⚠️ No structured data detected.")
            else:
                df = pd.DataFrame(data)
                st.success(f"✅ Extraction complete — {len(df)} valid records found.")
                st.dataframe(df, use_container_width=True)

                # Totals
                try:
                    total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                    total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                    net = round(total_debit - total_credit, 2)
                    st.markdown(f"**💰 Total Debit:** {total_debit:,.2f} | **Total Credit:** {total_credit:,.2f} | **Net:** {net:,.2f}")
                except:
                    pass

                st.download_button(
                    "⬇️ Download Excel",
                    data=to_excel_bytes(data),
                    file_name="vendor_statement_debe_haber.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("Please upload a vendor statement PDF to begin.")
