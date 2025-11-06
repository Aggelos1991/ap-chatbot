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
st.title("🦅 DataFalcon Pro — Hybrid GPT + OCR Extractor (Column-Aware)")

# === Load environment ===
try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None
PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"

# ==========================================================
# CORE REGEX + NORMALIZERS
# ==========================================================
RE_HEADER = re.compile(
    r"(?i)fecha.*asiento.*diario.*cuenta.*empresa.*ref.*cuenta.*anal|fecha.*asiento.*diario.*cuenta.*num.*debe.*haber.*saldo"
)
RE_SPLIT = re.compile(r"\s{2,}")  # split columns by 2+ spaces
RE_SKIP_ROW = re.compile(r"(?i)\b(asiento|diario|apertura|regularizaci[oó]n|saldo\s+anterior|sumas\s+y\s+saldos|total(?:es)?)\b")
RE_ACCOUNT_43 = re.compile(r"^43\d{6,}$")
RE_DOC_PREFERRED = re.compile(r"(?i)\b((?:A|AB|AC|FV|FAC|FA|CO)\s?\d{3,}|\d{5,})\b")
RE_DOC_HINT = re.compile(
    r"(?i)\b(num\.?|n[úu]m\.?|número|numero|documento|doc\.?|factura|fac|fv|co|ab|"
    r"αριθ\.?|αριθμός|παραστατικ|τιμολ[οό]γιο|παρ\.)\s*[:\-#]?\s*([A-Z]{0,3}\s?\d{3,}|\d{5,})"
)
RE_PAYMENT = re.compile(r"(?i)\b(orden\s+de\s+cobro|banco|bnki|pago|transferencia|remesa|cobro|trf|bank)\b")
RE_CREDITNOTE = re.compile(r"(?i)\b(reversi[oó]n|abono|nota\s*de\s*cr[eé]dito|cr[eé]dit|descuento|π[ίι]στωση)\b")

def normalize_number(value):
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    s = s.replace(" ", "")
    # handle (1.234,56) negatives
    neg = s.startswith("(") and s.endswith(")")
    if neg: s = s[1:-1]
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        v = round(float(s), 2)
        if neg: v = -v
        return v
    except:
        return ""

def prefer_doc_token(text):
    """
    Return best document-like token from text:
    - Prefer (A|AB|AC|FV|FAC|FA|CO)+digits
    - Else any ≥5 digits
    """
    if not text:
        return ""
    m = RE_DOC_PREFERRED.search(text.replace(" ", ""))
    if m:
        return re.sub(r"\s+", "", m.group(1))
    m2 = re.search(r"\d{5,}", text)
    return m2.group(0) if m2 else ""

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
# COLUMN-AWARE PARSER
# ==========================================================
COLUMN_ALIASES = {
    "fecha": "Date",
    "asiento": None,
    "diario": None,
    "cuenta": "Cuenta",
    "empresa": None,
    "ref": "Ref",
    "ref - etiqueta": "Ref",
    "etiqueta": "Ref",
    "cuenta analítica": None,
    "cuenta analitica": None,
    "num.": "Num",
    "num": "Num",
    "nº": "Num",
    "debe": "Debit",
    "haber": "Credit",
    "saldo acumulado": "Balance",
    "saldo": "Balance",
}

def detect_header(parts):
    """Return list of normalized names mapped to our schema or None if not a header."""
    joined = " ".join(parts)
    if not RE_HEADER.search(joined):
        return None
    norm = []
    for p in parts:
        k = p.strip().lower()
        # map by startswith for safety
        mapped = None
        for alias, target in COLUMN_ALIASES.items():
            if k.startswith(alias):
                mapped = target if target else alias  # keep original for ignored cols
                break
        norm.append(mapped if mapped else k)
    return norm

def parse_rows_columnar(lines):
    """
    Parse using detected header and column positions.
    Returns list of dict rows with Date, Alternative Document, Reason, Debit, Credit, Balance.
    """
    records = []
    header_map = None  # list of normalized labels per index
    for raw in lines:
        if RE_SKIP_ROW.search(raw):
            continue
        parts = RE_SPLIT.split(raw)
        # header detection
        hdr = detect_header(parts)
        if hdr:
            header_map = hdr
            continue
        if not header_map:
            continue  # we only parse after a header

        # if the line has fewer parts than header, skip (wrapped/continuation lines)
        if len(parts) < len(header_map) - 2:  # allow last column missing by OCR
            continue

        row = {}
        # map columns by index
        for idx, val in enumerate(parts[:len(header_map)]):
            label = header_map[idx]
            if not label:  # ignored columns
                continue
            v = val.strip()
            if label == "Date":
                row["Date"] = v
            elif label == "Cuenta":
                row["Cuenta"] = v
            elif label == "Ref":
                row["Ref"] = v
            elif label == "Num":
                row["Num"] = v
            elif label == "Debit":
                row["Debit"] = normalize_number(v)
            elif label == "Credit":
                row["Credit"] = normalize_number(v)
            elif label == "Balance":
                row["Balance"] = normalize_number(v)

        # derive Alternative Document
        alt_doc = ""
        # Prefer explicit Num
        if row.get("Num"):
            # never take account 43... from Num (rare, but safe)
            if not RE_ACCOUNT_43.match(row["Num"].replace(" ", "")):
                alt_doc = prefer_doc_token(row["Num"])
        # else search in Ref/Etiqueta
        if not alt_doc and row.get("Ref"):
            # accept patterns like "por factura 12345"
            m = RE_DOC_HINT.search(row["Ref"])
            if m:
                alt_doc = prefer_doc_token(m.group(0))
            if not alt_doc:
                alt_doc = prefer_doc_token(row["Ref"])

        if not alt_doc:
            # last resort: scan whole line, but NEVER take 43xxxxx
            tmp = prefer_doc_token(raw)
            if tmp and not RE_ACCOUNT_43.match(tmp.replace(" ", "")):
                alt_doc = tmp

        # determine reason
        text_for_reason = " ".join([raw, row.get("Ref", ""), row.get("Num", "")])
        if RE_PAYMENT.search(text_for_reason):
            reason = "Payment"
        elif RE_CREDITNOTE.search(text_for_reason):
            reason = "Credit Note"
        else:
            reason = "Invoice"

        # never accept Cuenta 43... as document
        if alt_doc and RE_ACCOUNT_43.match(alt_doc.replace(" ", "")):
            alt_doc = ""

        # amounts sanity + side repair
        debit_val = row.get("Debit", "")
        credit_val = row.get("Credit", "")
        balance_val = row.get("Balance", "")

        # if both missing but line clearly has numbers, try to recover last two numeric tokens
        if debit_val == "" and credit_val == "":
            nums = re.findall(r"[\-\(]?\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\)?", raw)
            if len(nums) >= 2:
                debit_val = normalize_number(nums[-2])
                credit_val = normalize_number(nums[-1])

        # side correction
        if reason == "Payment":
            if debit_val and not credit_val:
                credit_val, debit_val = debit_val, 0
        else:  # Invoice or Credit Note → left side (Debit)
            if credit_val and not debit_val:
                debit_val, credit_val = credit_val, 0

        # final validation
        has_amount = (debit_val not in ("", None) and float(debit_val) != 0.0) or \
                     (credit_val not in ("", None) and float(credit_val) != 0.0)
        if not has_amount:
            continue
        if not alt_doc:
            # keep row but mark unknown doc if you prefer; here we skip to keep it clean
            # (toggle if you want to keep unknowns)
            continue

        records.append({
            "Alternative Document": alt_doc,
            "Date": row.get("Date", ""),
            "Reason": reason,
            "Debit": debit_val if debit_val != "" else 0,
            "Credit": credit_val if credit_val != "" else 0,
            "Balance": balance_val if balance_val != "" else ""
        })

    return records

# ==========================================================
# GPT FALLBACK (OPTIONAL; OFF BY DEFAULT)
# ==========================================================
def gpt_repair_ambiguous(lines, existing_records):
    """
    Only attempts to fill missed rows (no doc) by scanning with GPT.
    Lines already parsed are not touched.
    """
    if not client:
        return existing_records

    taken_docs = {r["Alternative Document"] for r in existing_records}
    joined = "\n".join(f"{i+1}. {ln}" for i, ln in enumerate(lines))

    prompt = f"""
You are repairing a vendor statement extraction. The text is numbered lines.

Extract ONLY rows that have:
- A real document reference (Num./Factura/Doc/Concepto), NOT account codes like 43xxxxxx
- A valid amount (two columns: Debe and Haber; 'Saldo acumulado' is the running balance)

Return a pure JSON array of objects with:
- "Line": integer line index (1-based)
- "Alternative Document": code like A741387, AB0718, FV12345, FAC2345, CO1234, or any code with ≥5 digits (NEVER 43xxxxxx)
- "Date": date string on that row
- "Reason": "Invoice" | "Payment" | "Credit Note"
- "Debit": number
- "Credit": number
- "Balance": number (Saldo acumulado) if visible, else empty

Rules:
- Payment if text contains: "Orden de cobro", "Banco", "BNKI", "Pago", "Transferencia", "Remesa", "Cobro".
- Credit Note if: "Reversión", "Abono", "Nota de crédito", "Crédit", "Πίστωση".
- Else Invoice.
- Do NOT invent values. Skip if unsure.
- NEVER output an Alternative Document that starts with 43 and digits (account code).
- Return ONLY JSON.

Text:
{joined}
""".strip()

    out = []
    for model in [PRIMARY_MODEL, BACKUP_MODEL]:
        try:
            resp = client.chat.completions.create(
                model=model, temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )
            raw = resp.choices[0].message.content.strip()
            m = re.search(r'\[.*\]', raw, re.DOTALL)
            if not m:
                continue
            out = json.loads(m.group(0))
            break
        except Exception as e:
            st.warning(f"GPT fallback error ({model}): {e}")
            continue

    repaired = existing_records[:]
    for r in out:
        alt = str(r.get("Alternative Document", "")).strip()
        if not alt or alt in taken_docs:  # skip duplicates or empties
            continue
        if RE_ACCOUNT_43.match(alt.replace(" ", "")):
            continue

        debit_val = normalize_number(r.get("Debit", ""))
        credit_val = normalize_number(r.get("Credit", ""))
        balance_val = normalize_number(r.get("Balance", ""))

        reason = str(r.get("Reason", "")).strip().title()
        text_for_reason = str(r)

        if RE_PAYMENT.search(text_for_reason):
            reason = "Payment"
            if debit_val and not credit_val:
                credit_val, debit_val = debit_val, 0
        elif RE_CREDITNOTE.search(text_for_reason):
            reason = "Credit Note"
            if credit_val and not debit_val:
                debit_val, credit_val = credit_val, 0
        else:
            reason = "Invoice"
            if credit_val and not debit_val:
                debit_val, credit_val = credit_val, 0

        has_amount = (debit_val not in ("", None) and float(debit_val) != 0.0) or \
                     (credit_val not in ("", None) and float(credit_val) != 0.0)
        if not has_amount:
            continue

        repaired.append({
            "Alternative Document": alt,
            "Date": str(r.get("Date", "")).strip(),
            "Reason": reason,
            "Debit": debit_val if debit_val != "" else 0,
            "Credit": credit_val if credit_val != "" else 0,
            "Balance": balance_val if balance_val != "" else ""
        })
    return repaired

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
# UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])
use_gpt_fallback = st.checkbox("Enable GPT fallback for ambiguous rows (optional)", value=False)

if uploaded_pdf:
    with st.spinner("📄 Extracting text + running OCR fallback..."):
        lines, ocr_pages = extract_text_with_ocr(uploaded_pdf)

    if not lines:
        st.error("❌ No text detected. Ensure Tesseract (spa/eng/ell) is installed.")
    else:
        st.success(f"✅ Found {len(lines)} lines of text.")
        if ocr_pages:
            st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")
        st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

        if st.button("🔎 Parse Statement", type="primary"):
            with st.spinner("Parsing with column-aware engine..."):
                base_records = parse_rows_columnar(lines)
                records = base_records

                if use_gpt_fallback:
                    with st.spinner("Running GPT fallback for missed rows..."):
                        records = gpt_repair_ambiguous(lines, base_records)

            if records:
                df = pd.DataFrame(records)
                st.success(f"✅ Extraction complete — {len(df)} records.")
                st.dataframe(df, use_container_width=True, hide_index=True)

                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                valid_balances = df["Balance"].apply(pd.to_numeric, errors="coerce").dropna()
                final_balance = valid_balances.iloc[-1] if not valid_balances.empty else total_debit - total_credit

                c1, c2, c3 = st.columns(3)
                c1.metric("💰 Total Debit (Debe)", f"{total_debit:,.2f}")
                c2.metric("💳 Total Credit (Haber)", f"{total_credit:,.2f}")
                c3.metric("📊 Final Balance (Saldo)", f"{final_balance:,.2f}")

                st.download_button(
                    "⬇️ Download Excel",
                    data=to_excel_bytes(records),
                    file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ No valid transaction rows found. Try enabling GPT fallback.")
else:
    st.info("Upload a PDF to begin.")
