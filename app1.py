import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT+OCR Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Hybrid GPT + OCR Extractor")

# === Load environment ===
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
# OCR EXTRACTION
# ==========================================================
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
                    if clean:
                        all_lines.append(clean)
            else:
                ocr_pages.append(i)
                try:
                    img = convert_from_bytes(pdf_bytes, dpi=250, first_page=i, last_page=i)[0]
                    ocr_text = pytesseract.image_to_string(img, lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean:
                            all_lines.append(clean)
                except Exception as e:
                    st.warning(f"OCR skipped for page {i}: {e}")
    return all_lines, ocr_pages

# ==========================================================
# UTILITIES
# ==========================================================
NUM_TOKEN = re.compile(
    r"""(?ix)
    (?:
        (?:num\.?|n[úu]m\.?|número|numero|documento|doc\.?|factura|fac|fv|co|ab|
         αριθ|παραστατικ|τιμολόγιο|τιμολογιο|παρ\.)
        \s*[:\-#]?\s*
    )?
    ([A-Z]{0,3}\s?\d{3,}|\d{5,})
    """
)

PREFERRED_PATTERN = re.compile(r"(?i)\b((F|FV|CO|AB|FAC|FA)\d{3,}|\d{5,})\b")

SKIP_ROW = re.compile(r"(?i)\b(asiento|diario|apertura|regularizaci|saldo\s+anterior|sumas\s+y\s+saldos)\b")
SKIP_ALT = re.compile(r"(?i)\b(asiento|diario|apertura|regularizaci)\b")
GENERIC_REF = re.compile(r"(?i)^(pago|remesa|transfer|trf|bank)$")

def normalize_number(value):
    if value is None:
        return ""
    s = str(value).strip().replace(" ", "")
    if not s:
        return ""
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-()]", "", s)
    # Handle (123,45) negatives
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return round(float(s), 2)
    except:
        return ""

def parse_gpt_response(content, batch_num):
    json_match = re.search(r'\[.*\]', content, re.DOTALL)
    if not json_match:
        st.warning(f"⚠️ Batch {batch_num}: No JSON found.\n{content[:200]}")
        return []
    try:
        return json.loads(json_match.group(0))
    except json.JSONDecodeError as e:
        st.warning(f"⚠️ Batch {batch_num}: JSON decode error → {e}")
        return []

# ----------------------------------------------------------
# Local candidate scan per line (used for repair)
# ----------------------------------------------------------
def scan_candidates_per_line(lines):
    """
    Returns list of lists: candidates[idx] -> list of doc refs found on that line.
    """
    cands = []
    for ln in lines:
        hits = []
        for m in NUM_TOKEN.finditer(ln):
            token = m.group(1).strip()
            if token:
                # prefer compact form, strip spaces like "FV 12345" -> "FV12345"
                token = re.sub(r"\s+", "", token)
                hits.append(token)
        cands.append(hits)
    return cands

def best_candidate_from_neighbors(candidates, line_idx, radius=2):
    """
    Look at line_idx and neighbors ±radius to pick a preferred pattern first,
    otherwise the first token found.
    line_idx is 0-based.
    """
    n = len(candidates)
    pool = []
    for j in range(max(0, line_idx - radius), min(n, line_idx + radius + 1)):
        pool.extend(candidates[j])
    # Prefer pattern like FV/FAC/CO/AB/FA999...
    for tok in pool:
        if PREFERRED_PATTERN.search(tok):
            return PREFERRED_PATTERN.search(tok).group(1)
    # fallback: any ≥5 digits token
    for tok in pool:
        if re.search(r"\d{5,}", tok):
            return tok
    return ""

# ==========================================================
# GPT EXTRACTION  (HARDENED + CONTEXT REPAIR)
# ==========================================================
def extract_with_gpt(all_lines):
    BATCH_SIZE = 60
    all_records = []

    for batch_start in range(0, len(all_lines), BATCH_SIZE):
        batch_lines = all_lines[batch_start:batch_start + BATCH_SIZE]
        # pre-scan local candidates for this batch
        batch_candidates = scan_candidates_per_line(batch_lines)

        # number the lines for GPT and keep exactly same text
        numbered = "\n".join(f"{idx+1}. {txt}" for idx, txt in enumerate(batch_lines))

        prompt = f"""
You are a multilingual financial data extractor specialized in vendor statements (Spanish / Greek / English).

TASK:
Read the numbered lines below and extract ONLY real transaction lines:
- Invoice (Factura), Credit Note (Abono/Nota de crédito), Payment (Pago/Transferencia/Remesa/Cobro).
IGNORE accounting or ledger lines such as: "Asiento", "Diario", "Apertura", "Regularización", "Saldo anterior", totals, or summaries.

OUTPUT:
Return a pure JSON array. Each object MUST have:
- "Line" (integer) → the exact line number from the provided text.
- "Alternative Document" → the true invoice/credit/payment reference (from Num./Documento/Factura or from Concepto/comentarios like "por factura 12345").
- "Date" → the date on the same or nearby line (string).
- "Reason" ∈ ["Invoice","Payment","Credit Note"].
- "Debit"
- "Credit"
- "Balance"

RULES:
- If a line mentions "Pago", "Cobro", "Transferencia", "Remesa" ⇒ Reason = "Payment".
- If "Abono", "Nota de crédito", "Crédit", "Descuento", "Πίστωση" ⇒ Reason = "Credit Note".
- Else assume "Invoice".
- Do NOT output entries that are only "Asiento/Diario/Apertura/Regularización/Saldo anterior" (skip them).
- The document reference should look like: (F|FV|CO|AB|FA|FAC) + digits OR any code with ≥5 consecutive digits.
- Do NOT invent fields. If uncertain, skip that line.
- Return ONLY the JSON.

TEXT:
{numbered}
""".strip()

        data = []
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                content = response.choices[0].message.content.strip()
                if batch_start == 0:
                    st.text_area(f"🧠 GPT Response (Batch 1 – {model})", content, height=250, key=f"debug_{model}")
                data = parse_gpt_response(content, batch_start // BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error ({model}): {e}")

        if not data:
            continue

        # === Post-processing + Context Repair ===
        for row in data:
            # 1) Line mapping
            try:
                line_num = int(row.get("Line", 0))
            except:
                line_num = 0
            if not (1 <= line_num <= len(batch_lines)):
                continue
            line_idx = line_num - 1
            raw_line = batch_lines[line_idx]

            # 2) Skip obvious ledger lines
            if SKIP_ROW.search(raw_line):
                continue

            # 3) Extract fields
            alt_doc = str(row.get("Alternative Document", "")).strip()
            reason = str(row.get("Reason", "")).strip().title()
            date_str = str(row.get("Date", "")).strip()

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            balance_val = normalize_number(row.get("Balance", ""))

            # 4) Clean/validate Alternative Document
            if alt_doc:
                if SKIP_ALT.search(alt_doc) or GENERIC_REF.match(alt_doc):
                    alt_doc = ""
                else:
                    m = PREFERRED_PATTERN.search(alt_doc)
                    if m:
                        alt_doc = m.group(1)
                    else:
                        # maybe ≥5 digits token
                        m2 = re.search(r"\d{5,}", alt_doc)
                        alt_doc = m2.group(0) if m2 else ""

            # 5) Context repair: if still empty/invalid, pull from local candidates (same line ±2)
            if not alt_doc:
                alt_doc = best_candidate_from_neighbors(batch_candidates, line_idx, radius=2)
                if not alt_doc:
                    # final safety: try scanning the line raw
                    m3 = NUM_TOKEN.search(raw_line)
                    if m3:
                        alt_doc = re.sub(r"\s+", "", m3.group(1))

            if not alt_doc:
                # cannot trust this row without a real doc reference
                continue

            # 6) Reason correction from row/line context (robust)
            all_text = " ".join([str(row), raw_line])
            if re.search(r"(?i)pago|cobro|transfer|remesa|trf|bank", all_text):
                reason = "Payment"
            elif re.search(r"(?i)abono|nota\s*de\s*cr[eé]dito|cr[eé]dit|descuento|πίστωση", all_text):
                reason = "Credit Note"
            else:
                reason = "Invoice"

            # 7) Fix Debit/Credit placement
            if reason == "Payment":
                # Payments should hit Credit
                if debit_val and not credit_val:
                    credit_val, debit_val = debit_val, 0
            elif reason in ["Invoice", "Credit Note"]:
                # Invoices/Credit notes should hit Debit
                if credit_val and not debit_val:
                    debit_val, credit_val = credit_val, 0

            # 8) Skip blanks/zeros
            if (debit_val == "" or float(debit_val) == 0.0) and (credit_val == "" or float(credit_val) == 0.0):
                # if both empty or zero, ignore noisy rows
                continue

            # 9) Append
            all_records.append({
                "Alternative Document": alt_doc,
                "Date": date_str,
                "Reason": reason,
                "Debit": debit_val,
                "Credit": credit_val,
                "Balance": balance_val
            })

    return all_records

# ==========================================================
# EXPORT TO EXCEL
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT INTERFACE
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text + running OCR fallback..."):
        lines, ocr_pages = extract_text_with_ocr(uploaded_pdf)

    if not lines:
        st.error("❌ No text detected. Check that Tesseract OCR is installed and language packs (spa, ell, eng) are available.")
    else:
        st.success(f"✅ Found {len(lines)} lines of text!")
        if ocr_pages:
            st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")
        st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

        if st.button("🤖 Run Hybrid Extraction", type="primary"):
            with st.spinner("Analyzing with GPT + context repair..."):
                data = extract_with_gpt(lines)

            if data:
                df = pd.DataFrame(data)
                st.success(f"✅ Extraction complete — {len(df)} valid records found!")
                st.dataframe(df, use_container_width=True, hide_index=True)

                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                valid_balances = df["Balance"].apply(pd.to_numeric, errors="coerce").dropna()
                final_balance = valid_balances.iloc[-1] if not valid_balances.empty else total_debit - total_credit

                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("📊 Final Balance", f"{final_balance:,.2f}")

                st.download_button(
                    "⬇️ Download Excel",
                    data=to_excel_bytes(data),
                    file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ No structured data found in GPT output.")
else:
    st.info("Upload a PDF to begin.")
