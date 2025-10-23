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
    """Normalize decimals like 1.234,56 → 1234.56 and handle negatives"""
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
        num = float(s)
        return round(num, 2)
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
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]?\d{0,2}", line):
                    all_lines.append(" ".join(line.split()))
    return all_lines

# ==========================================================
# GPT EXTRACTOR — Enhanced Document Number Detection + Negatives
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to detect Debit (DEBE) and Credit (HABER) from vendor statements."""
    BATCH_SIZE = 150
    all_records = []
    
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        
        prompt = (
            "You are an expert accountant fluent in SPANISH, GREEK, and accounting terminology.\n"
            "You are reading extracted lines from a vendor statement (bank statement, AP statement, etc.).\n\n"
            "## DOCUMENT NUMBER DETECTION - CRITICAL\n"
            "Find document numbers in these formats/labels (prioritize in this order):\n"
            "1. Spanish: Nº, Num, Número, Documento, Factura, Fra, Ref, Referencia, Fact, Fatura\n"
            "2. Greek: Τιμολόγιο (Timologio), Αριθμός (Arithmos), Αρ., Νο., Παραστατικό, Τ/Λ, ΤΛ\n"
            "3. Common: Invoice #, DOC, ID, RefNo\n"
            "4. Numbers alone: 1-3 digits followed by dashes/dots or 6+ digits (e.g., 123, 123-45, 2024/001)\n\n"
            "## TRANSACTION COLUMNS\n"
            "- DEBE: Debit/Invoice amount (Fra. emitida, Cargo)\n"
            "- HABER: Credit/Payment amount (Cobro, Pago, Abono)\n"
            "- SALDO: Running balance (ignore for extraction)\n"
            "- CONCEPTO: Description\n\n"
            "## NEGATIVE NUMBER RULE - IMPORTANT\n"
            "- Negative in DEBE → Move to CREDIT and classify as 'Credit Note'\n"
            "- Negative in HABER → Move to DEBIT and classify as 'Invoice'\n\n"
            "## CLASSIFICATION RULES\n"
            "1. Invoice: DEBE > 0 OR contains 'Fra', 'Factura', 'Τιμολόγιο', 'emitida'\n"
            "2. Payment: HABER > 0 OR contains 'Cobro', 'Pago', 'Είσπραξη', 'Επιταγή'\n"
            "3. Credit Note: Contains 'NC', 'Nota Credito', 'Ακυρωτικό', 'Πιστωτικό', 'Abono' OR NEGATIVE DEBE\n\n"
            "## OUTPUT FORMAT - EXACTLY\n"
            "Return ONLY a valid JSON array. Each object:\n"
            """[
  {
    "Alternative Document": "EXACT document number found (e.g. 'FRA-2024-001', '12345', 'ΤΛ 67890')",
    "Date": "dd/mm/yyyy OR dd/mm/yy OR empty string",
    "Reason": "Invoice|Payment|Credit Note",
    "Debit": "numeric value OR empty string",
    "Credit": "numeric value OR empty string",
    "Description": "short description of transaction"
  }
]"""
            "\n\n## FILTERS - EXCLUDE THESE\n"
            "- Lines with 'concil', 'conciliacion', 'reconcil', 'reconciliacion' (case insensitive)\n"
            "- Summary totals: 'Total', 'Saldo', 'Apertura', 'Cierre', 'IVA', 'Base Imponible'\n"
            "- Empty document numbers\n\n"
            f"Lines to analyze:\n\"\"\"{text_block}\"\"\""
        )
        
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1
            )
            content = response.choices[0].message.content.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            if not json_match:
                st.warning(f"⚠️ No valid JSON in batch {i//BATCH_SIZE + 1}")
                continue
            data = json.loads(json_match.group(0))
        except Exception as e:
            st.warning(f"⚠️ GPT failed on batch {i//BATCH_SIZE + 1}: {e}")
            continue
        
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            
            # 🚫 Enhanced exclusion filter
            exclude_patterns = [
                r"concil", r"conciliacion", r"reconciliacion", r"reconcil",
                r"saldo\b", r"total\b", r"iva\b", r"apertura", r"cierre"
            ]
            exclude_text = alt_doc + str(row.get("Description", ""))
            if any(re.search(pattern, exclude_text, re.IGNORECASE) for pattern in exclude_patterns):
                continue
            
            # Validate document number (must have digits)
            if not re.search(r"\d", alt_doc):
                continue
                
            debit_raw = row.get("Debit", "")
            credit_raw = row.get("Credit", "")
            
            debit_val = normalize_number(debit_raw)
            credit_val = normalize_number(credit_raw)
            
            # 🆕 Handle negative numbers (convert DEBE negative → Credit Note)
            reason = row.get("Reason", "").strip()
            if debit_val and float(debit_val) < 0:
                credit_val = abs(float(debit_val))
                debit_val = ""
                reason = "Credit Note"
            elif credit_val and float(credit_val) < 0:
                debit_val = abs(float(credit_val))
                credit_val = ""
                reason = "Invoice"
            
            # Final classification based on amounts
            if debit_val and float(debit_val) > 0 and reason.lower() != "credit note":
                reason = "Invoice"
            elif credit_val and float(credit_val) > 0:
                reason = "Payment"
            
            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": reason,
                "Debit": debit_val,
                "Credit": credit_val,
                "Description": str(row.get("Description", "")).strip()
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
        
        col1, col2 = st.columns([3,1])
        with col1:
            if st.button("🤖 Run Enhanced Hybrid Extraction", type="primary"):
                with st.spinner("🔍 Analyzing with GPT-4o-mini (Enhanced Doc + Negatives)..."):
                    data = extract_with_gpt(lines)
                
                if not data:
                    st.warning("⚠️ No structured data detected.")
                else:
                    df = pd.DataFrame(data)
                    st.success(f"✅ Extraction complete — {len(df)} valid records found!")
                    
                    # Enhanced display
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    
                    # Summary metrics
                    try:
                        df_num = df.copy()
                        df_num["Debit"] = pd.to_numeric(df_num["Debit"], errors="coerce")
                        df_num["Credit"] = pd.to_numeric(df_num["Credit"], errors="coerce")
                        
                        total_debit = df_num["Debit"].sum()
                        total_credit = df_num["Credit"].sum()
                        net = round(total_debit - total_credit, 2)
                        
                        col_a, col_b, col_c, col_d = st.columns(4)
                        with col_a:
                            st.metric("💰 Total Debit", f"{total_debit:,.2f}")
                        with col_b:
                            st.metric("💳 Total Credit", f"{total_credit:,.2f}")
                        with col_c:
                            st.metric("⚖️ Net Balance", f"{net:,.2f}")
                        with col_d:
                            st.metric("📊 Records", len(df))
                            
                    except Exception as e:
                        st.error(f"Summary calculation error: {e}")
                    
                    # Download
                    excel_data = to_excel_bytes(data)
                    st.download_button(
                        "⬇️ Download Excel",
                        data=excel_data,
                        file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        # Debug info
        with col2:
            st.markdown("### 📈 Stats")
            st.metric("Lines analyzed", len(lines))
            if st.button("🔍 Show sample GPT prompt"):
                st.info("✅ Enhanced prompt includes Spanish/Greek doc detection + negative handling!")
else:
    st.info("👆 Please upload a vendor statement PDF to begin extraction.")
    
    st.markdown("""
    ## ✨ **Enhanced Features**
    - **Document Detection**: Spanish (Factura, Nº, Fra) + Greek (Τιμολόγιο, Αρ.)
    - **🆕 Negative Handling**: DEBE(-100) → Credit Note +100
    - **Smart Classification**: Auto-detects Invoice/Payment/Credit Note
    - **Multi-language**: Spanish, Greek, English accounting terms
    - **Validation**: Filters out reconciliation/summary lines
    """)
