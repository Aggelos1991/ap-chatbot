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
PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"


# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(value):
    """Normalize decimals like 1.234,56 → 1234.56."""
    if not value:
        return ""
    s = str(value).strip().replace(" ", "")

    if "," in s and "." in s:
        # European format 1.234,56 → 1234.56
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
# PDF TEXT EXTRACTION — OCR REMOVED COMPLETELY
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    """
    Extract ALL pure text lines from every page using pdfplumber only.
    No OCR fallback.
    """
    all_lines = []

    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            if not text:
                # If a page has no extractable text, just skip it.
                continue

            for line in text.split("\n"):
                clean_line = " ".join(line.split())

                if not clean_line:
                    continue

                # Ignore Saldo lines
                if re.search(r"\bsaldo\b", clean_line, re.IGNORECASE):
                    continue

                all_lines.append(clean_line)

    return all_lines


# ==========================================================
# GPT RESPONSE PARSER
# ==========================================================
def parse_gpt_response(content, batch_num):
    json_match = re.search(r'\[.*\]', content, re.DOTALL)
    if not json_match:
        st.warning(f"⚠️ Batch {batch_num}: No JSON found. First 300 chars:\n{content[:300]}")
        return []

    try:
        return json.loads(json_match.group(0))
    except json.JSONDecodeError as e:
        st.warning(f"⚠️ Batch {batch_num}: JSON decode error → {e}")
        return []


# ==========================================================
# GPT EXTRACTOR — CORE LOGIC
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are a financial data extractor specialized in Spanish and Greek vendor statements.

Each line may contain:
- Fecha / Date
- Documento / N° DOC / Αρ. Τιμολογίου
- Concepto / Περιγραφή
- DEBE (Invoice)
- HABER (Payment or Credit Note)
- TOTAL / TOTALES when DEBE/HABER missing

RULES:
1. Ignore: Asiento, Saldo, IVA, Total Saldo.
2. Exclude anything like "Código IC N".
3. Extract invoice-like codes when DOC missing.
4. Detect reason automatically:
   - Payments → Cobro, Pago, Transferencia, Remesa, Bank, Trf
   - Credit Notes → Abono, Nota de crédito, Crédito, Πίστωση
   - Invoices → Fra., Factura, Τιμολόγιο
5. If only TOTAL exists → treat as Debit.
6. Output JSON only. No explanations.

OUTPUT FORMAT:
[
  {{
    "Alternative Document": "",
    "Concepto": "",
    "Date": "",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "",
    "Credit": ""
  }}
]

TEXT:
{text_block}
"""

        # Try primary model then backup
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.0
                )

                content = response.choices[0].message.content.strip()

                # Only show debug for first batch
                if i == 0:
                    st.text_area(f"🧠 GPT Response (Batch 1 – {model})", content, height=250)

                data = parse_gpt_response(content, i // BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"❌ GPT error with {model}: {e}")
                data = []

        if not data:
            continue

        # === CLEAN & NORMALIZE RECORDS ===
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()

            # Exclude Código IC lines
            if re.search(r"codigo\s*ic\s*n", alt_doc, re.IGNORECASE):
                continue

            if not alt_doc:
                continue

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            reason = row.get("Reason", "").strip()

            # Negative DEBE → Credit Note
            if debit_val != "" and credit_val == "":
                try:
                    if float(debit_val) < 0:
                        credit_val = round(abs(float(debit_val)), 2)
                        debit_val = ""
                        reason = "Credit Note"
                    else:
                        reason = "Invoice"
                except:
                    reason = "Invoice"

            elif credit_val != "" and debit_val == "":
                if re.search(r"abono|nota|crédit|descuento|πίστωση", str(row), re.IGNORECASE):
                    reason = "Credit Note"
                else:
                    reason = "Payment"

            elif debit_val == "" and credit_val == "":
                continue

            all_records.append({
                "Alternative Document": alt_doc,
                "Concepto": row.get("Concepto", "").strip(),
                "Date": row.get("Date", "").strip(),
                "Reason": reason,
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

    with st.spinner("📄 Extracting text..."):
        lines = extract_raw_lines(uploaded_pdf)

    st.success(f"✅ Extracted {len(lines)} lines (Saldo removed).")

    st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Analyzing with GPT..."):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)

            st.success(f"✅ Extraction complete — {len(df)} records found!")

            st.dataframe(
                df[
                    ["Alternative Document", "Date", "Concepto", "Reason", "Debit", "Credit"]
                ],
                use_container_width=True,
                hide_index=True
            )

            # Totals
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
            st.warning("⚠️ No structured data produced.")

else:
    st.info("Please upload a vendor statement PDF to begin.")
