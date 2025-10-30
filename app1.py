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
    """Extract ALL text lines from every page of the PDF."""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                clean_line = " ".join(line.split())
                if clean_line.strip():
                    all_lines.append(clean_line)
    return all_lines

# ==========================================================
# GPT EXTRACTOR — Enhanced with Comentario / Concepto logic
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to detect Debit (DEBE) and Credit (HABER) from vendor statements."""
    BATCH_SIZE = 100
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a financial data extractor specialized in Spanish vendor statements.

Each line contains:
- Fecha (Date)
- N° DOC or Documento (Document number)
- Comentario / Concepto / Descripción (text with context — may contain invoice or payment details)
- DEBE (Invoice amounts)
- HABER / CRÉDITO (Payments or credit notes)
- SALDO (running balance — IGNORE COMPLETELY)

Your task: extract all valid transactions and classify them precisely.

**CLASSIFICATION RULES:**
1. Never use "Asiento", "Saldo", "IVA" or "Total" as document or transaction lines — ignore them.
2. If "N° DOC" is empty, look for an invoice-like pattern in the "Comentario" text (examples: FRA 209, FAC1234, FACTURA 1775, INV-2024-01, etc.) and use that as Alternative Document.
3. Use **Comentario / Concepto / Descripción** text to detect the reason:
   - If it contains "Cobro", "Pago", "Transferencia", "Remesa", "Bank", "Trf", "Pagado" → it's a **Payment**.
   - If it contains "Abono", "Nota de crédito", "Crédito", "Descuento" → it's a **Credit Note**.
   - If it contains "Fra.", "Factura", "FRA", "Factura Proveedor" → it's an **Invoice**.
4. DEBE → always "Invoice".
5. HABER / CRÉDITO → "Payment" or "Credit Note" based on keywords.
6. If both DEBE and HABER appear, only one should be used — keep the correct side based on reason.
7. Never use SALDO values.
8. Leave blank if value is missing; never invent numbers.

**OUTPUT FORMAT (JSON only):**
[
  {{
    "Alternative Document": "...",
    "Date": "dd/mm/yy",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "DEBE amount or empty string",
    "Credit": "HABER amount or empty string"
  }}
]

Text to analyze:
{text_block}
"""

        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.0
            )
            content = response.choices[0].message.content.strip()

            if i == 0:
                st.text_area("GPT Response (Batch 1):", content, height=200, key="debug_1")

            json_match = re.search(r'\[.*\]', content, re.DOTALL)
            if not json_match:
                st.warning(f"No JSON found in batch {i//BATCH_SIZE + 1}")
                continue

            data = json.loads(json_match.group(0))

            for row in data:
                alt_doc = str(row.get("Alternative Document", "")).strip()
                if not alt_doc or re.search(r"(asiento|saldo|total|iva)", alt_doc, re.IGNORECASE):
                    continue

                debit_val = normalize_number(row.get("Debit", ""))
                credit_val = normalize_number(row.get("Credit", ""))
                reason = row.get("Reason", "").strip()

                # === SALDO / DOUBLE-SIDE CLEANUP LOGIC ===
                if debit_val and credit_val:
                    if reason.lower() in ["payment", "credit note"]:
                        debit_val = ""  # keep only credit
                    elif reason.lower() == "invoice":
                        credit_val = ""  # keep only debit
                    else:
                        # smaller one likely SALDO or duplicate
                        if abs(debit_val - credit_val) < 0.01 or min(debit_val, credit_val) / max(debit_val, credit_val) < 0.3:
                            if debit_val < credit_val:
                                debit_val = ""
                            else:
                                credit_val = ""

                # Enforce one-sided rule
                if debit_val and not credit_val:
                    reason = "Invoice"
                elif credit_val and not debit_val:
                    if re.search(r"abono|nota|crédit|descuento", str(row), re.IGNORECASE):
                        reason = "Credit Note"
                    else:
                        reason = "Payment"
                elif debit_val == "" and credit_val == "":
                    continue

                all_records.append({
                    "Alternative Document": alt_doc,
                    "Date": str(row.get("Date", "")).strip(),
                    "Reason": reason,
                    "Debit": debit_val,
                    "Credit": credit_val
                })

        except Exception as e:
            st.warning(f"GPT error batch {i//BATCH_SIZE + 1}: {e}")
            continue

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

    st.success(f"✅ Found {len(lines)} lines of text!")
    st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Analyzing with GPT-4o-mini..."):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success(f"✅ Extraction complete — {len(df)} valid records found!")
            st.dataframe(df, use_container_width=True, hide_index=True)

            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)

                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net", f"{net:,.2f}")
            except Exception as e:
                st.error(f"Totals error: {e}")

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No structured data detected. Check GPT response above.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
