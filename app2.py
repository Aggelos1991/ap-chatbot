import streamlit as st
import pandas as pd
import re

# ======================================
# CONFIGURATION
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ======================================
# HELPERS
# ======================================
def normalize_number(v):
    """Convert numeric strings like '1.234,56' or '1,234.56' safely to float."""
    if v is None or str(v).strip() == "":
        return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") == 1:
        # 1.234,56 -> 1234.56  OR  1,234.56 -> 1234.56
        if s.find(",") > s.find("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") == 1:
        # 123,45 -> 123.45
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        # 1.234.567 -> 1234567
        s = s.replace(".", "", s.count(".") - 1)
    try:
        return float(s)
    except:
        return 0.0


def normalize_columns(df: pd.DataFrame, tag: str) -> pd.DataFrame:
    """
    Map multilingual headers to unified names:
    - invoice_<tag>, credit_<tag>, debit_<tag>, reason_<tag>, cif_<tag>, date_<tag>
    If debit/credit missing, add zeros. Keeps all other columns.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    # Make sure columns are strings
    df = df.copy()
    df.columns = [str(c) for c in df.columns]

    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document",
            # Greek
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            # Greek
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            # Greek
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            # Greek
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            # Greek (safe)
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            # Greek
            "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού"
        ],
    }

    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    rename_map = {}

    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                # first hit wins
                if col not in rename_map.values():
                    rename_map[col] = f"{key}_{tag}"

    out = df.rename(columns=rename_map)

    # Ensure mandatory numeric columns
    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0

    # Ensure reason exists
    if f"reason_{tag}" not in out.columns:
        out[f"reason_{tag}"] = ""

    # Try to ensure invoice exists — if not mapped, pick a likely column or create blank
    if f"invoice_{tag}" not in out.columns:
        # Heuristic: first column that looks id-like
        likely = None
        for c in out.columns:
            if re.search(r"(invoice|fact|doc|ref|num|nº|αρ|τιμ|παραστα)", c.lower()):
                likely = c
                break
        if likely and likely not in [f"debit_{tag}", f"credit_{tag}", f"reason_{tag}"]:
            out[f"invoice_{tag}"] = out[likely].astype(str)
        else:
            out[f"invoice_{tag}"] = ""

    # Coerce key columns to string
    for c in [f"invoice_{tag}", f"reason_{tag}"]:
        out[c] = out[c].astype(str).str.strip()

    return out


def extract_digits(v) -> str:
    """Keep only digits, strip leading zeros."""
    return re.sub(r"\D", "", str(v or "")).lstrip("0")


def clean_invoice_code(v) -> str:
    """
    Normalize invoice strings for comparison:
    - drop common prefixes (inv, cn, τιμ, αρ, etc.)
    - remove year snippets (20xx)
    - strip non-alphanumerics
    - remove leading zeros
    - finally keep digits only (robust suffix/clean compare)
    """
    if not v:
        return ""
    s = str(v).strip().lower()
    s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no)\W*", "", s)
    s = re.sub(r"20\d{2}", "", s)              # remove obvious years
    s = re.sub(r"[^a-z0-9]", "", s)            # keep alnum
    s = re.sub(r"^0+", "", s)                  # trim leading zeros
    s = re.sub(r"[^\d]", "", s)                # keep only digits for final match logic
    return s


# ======================================
# DOC TYPE + AMOUNT LOGIC
# ======================================
PAYMENT_PATTERNS = [
    r"πληρωμ", r"απόδειξη\s*πληρωμ",
    r"payment", r"bank\s*transfer", r"trf",
    r"remesa", r"pago", r"transferencia"
]
INVOICE_EXCLUSIONS = ["τιμολόγιο", "invoice", "factura", "παραστατικό", "έξοδα", "expenses"]

CREDIT_WORDS = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση"]
INVOICE_WORDS = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]


def detect_erp_doc_type(row: pd.Series) -> str:
    reason = str(row.get("reason_erp", "") or "").lower()
    charge = normalize_number(row.get("debit_erp"))
    credit = normalize_number(row.get("credit_erp"))
    try:
        if any(re.search(p, reason) for p in PAYMENT_PATTERNS) and not any(x in reason for x in INVOICE_EXCLUSIONS):
            return "IGNORE"
    except Exception:
        pass

    if any(k in reason for k in CREDIT_WORDS):
        return "CN"
    if any(k in reason for k in INVOICE_WORDS) or credit > 0:
        return "INV"
    return "UNKNOWN"


def detect_vendor_doc_type(row: pd.Series) -> str:
    reason = str(row.get("reason_ven", "") or "").lower()
    debit = normalize_number(row.get("debit_ven"))
    credit = normalize_number(row.get("credit_ven"))
    try:
        if any(re.search(p, reason) for p in PAYMENT_PATTERNS) and not any(x in reason for x in INVOICE_EXCLUSIONS):
            return "IGNORE"
    except Exception:
        pass

    if any(k in reason for k in CREDIT_WORDS) or credit > 0:
        return "CN"
    if any(k in reason for k in INVOICE_WORDS) or debit > 0:
        return "INV"
    return "UNKNOWN"


def calc_erp_amount(row: pd.Series) -> float:
    doc = row.get("__doctype", "")
    charge = normalize_number(row.get("debit_erp"))
    credit = normalize_number(row.get("credit_erp"))
    if doc == "INV":
        return abs(credit)
    if doc == "CN":
        # CN on ERP typically shows as charge (debit) positive OR credit negative
        return -abs(charge if charge > 0 else credit)
    return 0.0


def calc_vendor_amount(row: pd.Series) -> float:
    doc = row.get("__doctype", "")
    debit = normalize_number(row.get("debit_ven"))
    credit = normalize_number(row.get("credit_ven"))
    if doc == "INV":
        return abs(debit)
    if doc == "CN":
        return -abs(credit if credit > 0 else debit)
    return 0.0


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    if erp_df is None or ven_df is None or erp_df.empty or ven_df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Detect types + amounts
    erp_df = erp_df.copy()
    ven_df = ven_df.copy()

    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)

    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # Merge multiple entries of same ERP invoice number (combine INV & CN to net)
    merged_rows = []
    if "invoice_erp" in erp_use.columns:
        grouped = erp_use.groupby("invoice_erp", dropna=False)
    else:
        # fallback: group by index block of 1
        grouped = [(None, erp_use)]

    for inv, group in grouped:
        if group.empty:
            continue

        # If 3+ entries, keep the last (latest)
        if len(group) >= 3:
            group = group.tail(1)

        inv_rows = group[group["__doctype"] == "INV"]
        cn_rows = group[group["__doctype"] == "CN"]

        if not inv_rows.empty and not cn_rows.empty:
            total_inv = inv_rows["__amt"].sum()
            total_cn = cn_rows["__amt"].sum()
            net = round(total_inv + total_cn, 2)
            base_row = inv_rows.iloc[-1].copy()
            base_row["__amt"] = net
            merged_rows.append(base_row)
        else:
            merged_rows.append(group.iloc[-1])

    erp_use = pd.DataFrame(merged_rows).reset_index(drop=True) if merged_rows else erp_use.reset_index(drop=True)

    # Matching
    matched = []
    used_vendor_rows = set()

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        e_digits = extract_digits(e_inv)
        e_code = clean_invoice_code(e_inv)

        best_take = None

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue

            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            v_digits = extract_digits(v_inv)
            v_code = clean_invoice_code(v_inv)

            diff = round(e_amt - v_amt, 2)
            amt_close = abs(diff) < 0.05
            same_type = (e.get("__doctype") == v.get("__doctype"))

            # similarity checks
            same_full = (e_inv == v_inv) and (e_inv != "")
            same_clean = (e_code != "" and e_code == v_code)
            len_diff = abs(len(e_code) - len(v_code))
            suffix_ok = (
                e_code != "" and v_code != "" and
                len(e_code) > 2 and len(v_code) > 2 and
                len_diff <= 2 and (e_code.endswith(v_code) or v_code.endswith(e_code))
            )

            # Acceptance rules
            take_it = False
            if same_type and same_full:
                take_it = True           # full code wins even if amount differs (we’ll mark Difference)
            elif same_type and (same_clean or suffix_ok) and amt_close:
                take_it = True           # relaxed code match + close amount

            if take_it:
                best_take = {
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match" if amt_close else "Difference"
                }
                best_v = v_idx
                break

        if best_take is not None:
            matched.append(best_take)
            used_vendor_rows.add(best_v)

    matched_df = pd.DataFrame(matched)

    # Build missing lists (based on invoice text)
    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_ven = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()

    if "invoice_ven" in ven_use.columns:
        missing_in_erp = ven_use[~ven_use["invoice_ven"].astype(str).isin(matched_ven)][["invoice_ven", "__amt"]].copy()
        missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    else:
        missing_in_erp = pd.DataFrame(columns=["Invoice", "Amount"])

    if "invoice_erp" in erp_use.columns:
        missing_in_vendor = erp_use[~erp_use["invoice_erp"].astype(str).isin(matched_erp)][["invoice_erp", "__amt"]].copy()
        missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    else:
        missing_in_vendor = pd.DataFrame(columns=["Invoice", "Amount"])

    # Force Invoice column to string
    for df in [matched_df, missing_in_erp, missing_in_vendor]:
        if not df.empty and "Invoice" in df.columns:
            df["Invoice"] = df["Invoice"].astype(str).str.strip()

    return matched_df, missing_in_erp, missing_in_vendor


# ======================================
# PAYMENT EXTRACTION
# ======================================
def extract_payments(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    keywords = [
        "pago", "pagos", "payment", "transfer", "transferencia", "bank", "trf",
        "remesa", "prepago", "ajuste", "πληρωμή", "μεταφορά", "τραπεζικό έμβασμα"
    ]

    def is_payment(x):
        return any(k in str(x).lower() for k in keywords)

    erp_pay = erp_df.copy()
    ven_pay = ven_df.copy()

    if "reason_erp" in erp_pay:
        erp_pay = erp_pay[erp_pay["reason_erp"].apply(is_payment)].copy()
    else:
        erp_pay = pd.DataFrame()

    if "reason_ven" in ven_pay:
        ven_pay = ven_pay[ven_pay["reason_ven"].apply(is_payment)].copy()
    else:
        ven_pay = pd.DataFrame()

    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_erp")) - normalize_number(r.get("credit_erp"))),
            axis=1
        )
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_ven")) - normalize_number(r.get("credit_ven"))),
            axis=1
        )

    matched_payments = []
    used_vendor = set()
    for _, e in erp_pay.iterrows():
        for v_idx, v in ven_pay.iterrows():
            if v_idx in used_vendor:
                continue
            diff = abs(e["Amount"] - v["Amount"])
            if diff < 0.05:
                matched_payments.append({
                    "ERP Reason": e.get("reason_erp", ""),
                    "Vendor Reason": v.get("reason_ven", ""),
                    "ERP Amount": round(float(e["Amount"]), 2),
                    "Vendor Amount": round(float(v["Amount"]), 2),
                    "Difference": round(diff, 2)
                })
                used_vendor.add(v_idx)
                break

    return erp_pay, ven_pay, pd.DataFrame(matched_payments)


# ======================================
# STREAMLIT UI
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        # Always read as strings to avoid 123.0 on invoice codes
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)

        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")

        with st.spinner("Reconciling invoices..."):
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
            erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)

        st.success("✅ Reconciliation complete")

        # ====== MATCHED / DIFFERENCES ======
        st.subheader("📊 Matched / Differences")
        if not matched.empty:
            # No .style here (Streamlit renders DataFrames directly)
            st.dataframe(matched, use_container_width=True)
        else:
            st.info("No matches found.")

        # ====== MISSING ======
        left, right = st.columns(2)
        with left:
            st.subheader("❌ Missing in ERP (present in Vendor)")
            if not erp_missing.empty:
                st.dataframe(erp_missing, use_container_width=True)
            else:
                st.success("✅ No missing invoices in ERP.")
        with right:
            st.subheader("❌ Missing in Vendor (present in ERP)")
            if not ven_missing.empty:
                st.dataframe(ven_missing, use_container_width=True)
            else:
                st.success("✅ No missing invoices in Vendor.")

        # ====== PAYMENTS ======
        st.subheader("🏦 Payment Transactions (Identified in both)")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**💼 ERP Payments**")
            if not erp_pay.empty:
                st.dataframe(erp_pay[["reason_erp", "debit_erp", "credit_erp", "Amount"]], use_container_width=True)
                st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR")
            else:
                st.info("No ERP payments found.")
        with col2:
            st.markdown("**🧾 Vendor Payments**")
            if not ven_pay.empty:
                st.dataframe(ven_pay[["reason_ven", "debit_ven", "credit_ven", "Amount"]], use_container_width=True)
                st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR")
            else:
                st.info("No Vendor payments found.")

        st.markdown("### ✅ Matched Payments")
        if not matched_pay.empty:
            st.dataframe(matched_pay, use_container_width=True)
            total_erp = matched_pay["ERP Amount"].sum()
            total_vendor = matched_pay["Vendor Amount"].sum()
            diff_total = abs(total_erp - total_vendor)
            st.markdown(f"**Total Matched ERP Payments:** {total_erp:,.2f} EUR")
            st.markdown(f"**Total Matched Vendor Payments:** {total_vendor:,.2f} EUR")
            st.markdown(f"**Difference Between ERP and Vendor Payments:** {diff_total:,.2f} EUR")
        else:
            st.info("No matching payments found.")

    except Exception as e:
        st.error("❌ Failed to process files. Check your columns and try again.")
        st.exception(e)
else:
    st.info("Upload both files to start: ERP export and Vendor statement (.xlsx).")
