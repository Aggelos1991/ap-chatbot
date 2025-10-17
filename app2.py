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
        if s.find(",") > s.find("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") == 1:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "", s.count(".") - 1)
    try:
        return float(s)
    except:
        return 0.0


def normalize_columns(df, tag):
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura",
            "alternative document", "αρ.", "αριθμός", "νούμερο", "no", "παραστατικό"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "importe haber", "valor haber", "πίστωση", "πιστωτικό"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales",
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "detalle", "razon",
            "observaciones", "comentario", "explicacion", "αιτιολογία", "περιγραφή"
        ],
        "date": [
            "date", "fecha", "fech", "data", "ημερομηνία", "ημ/νία"
        ],
    }

    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"

    out = df.rename(columns=rename_map)

    # Ensure debit/credit exist
    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0

    return out


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    # ---------- Detect doc types ----------
    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        payment_patterns = [
            r"^πληρωμ", r"^payment", r"^bank\s*transfer", r"^pago", r"^transferencia", r"^remesa", r"^trf"
        ]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in credit_words):
            return "CN"
        elif any(k in reason for k in invoice_words) or credit > 0:
            return "INV"
        return "UNKNOWN"

    def calc_erp_amount(row):
        doc = row.get("__doctype", "")
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if doc == "INV":
            return abs(credit)
        elif doc == "CN":
            return -abs(charge if charge > 0 else credit)
        return 0.0

    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        payment_words = ["pago", "payment", "transfer", "bank", "saldo", "trf", "πληρωμή", "τραπεζ"]
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in payment_words):
            return "IGNORE"
        elif any(k in reason for k in credit_words) or credit > 0:
            return "CN"
        elif any(k in reason for k in invoice_words) or debit > 0:
            return "INV"
        return "UNKNOWN"

    def calc_vendor_amount(row):
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        doc = row.get("__doctype", "")
        if doc == "INV":
            return abs(debit)
        elif doc == "CN":
            return -abs(credit if credit > 0 else debit)
        return 0.0

    # ---------- Apply classification ----------
    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # ---------- Cancelled invoices ----------
    def clean_invoice_code_simple(v):
        if not v:
            return ""
        s = str(v).strip().lower()
        s = re.sub(r"[\s./\-]+", "", s)
        s = re.sub(r"20\d{2}", "", s)
        s = re.sub(r"[^a-z0-9]", "", s)
        return s

    erp_use["__clean"] = erp_use["invoice_erp"].apply(clean_invoice_code_simple)
    ven_use["__clean"] = ven_use["invoice_ven"].apply(clean_invoice_code_simple)

    canceled_erp_df, canceled_ven_df = pd.DataFrame(), pd.DataFrame()

    for df, tag in [(erp_use, "erp"), (ven_use, "ven")]:
        mask = df.duplicated(subset=["__clean"], keep=False)
        canceled_codes, canceled_groups = set(), []
        for code, grp in df[mask].groupby("__clean"):
            if len(grp["__doctype"].unique()) >= 2:
                inv_amt = grp.loc[grp["__doctype"] == "INV", "__amt"].sum()
                cn_amt = grp.loc[grp["__doctype"] == "CN", "__amt"].sum()
                total = round(inv_amt + cn_amt, 2)
                if abs(total) < 0.05 and len(grp) == 2:
                    canceled_codes.add(code)
                    canceled_groups.append(grp)
        if canceled_groups:
            canceled_df = pd.concat(canceled_groups, ignore_index=True)
            if tag == "erp":
                canceled_erp_df = canceled_df
                erp_use = erp_use[~erp_use["__clean"].isin(canceled_codes)]
            else:
                canceled_ven_df = canceled_df
                ven_use = ven_use[~ven_use["__clean"].isin(canceled_codes)]

    # ---------- Matching logic ----------
    def extract_digits(v):
        if not v:
            return ""
        return re.sub(r"\D", "", str(v)).lstrip("0")

    def clean_invoice_code(v):
        if not v:
            return ""
        s = str(v).strip().lower()
        s = re.sub(r"[\s./\-]+", "", s)
        s = re.sub(r"20\d{2}", "", s)
        s = re.sub(r"[^a-z0-9]", "", s)
        s = re.sub(r"^0+", "", s)
        s = re.sub(r"[^\d]", "", s)
        return s

    for _, e in erp_use.iterrows():
        e_inv, e_amt = str(e.get("invoice_erp", "")).strip(), round(float(e["__amt"]), 2)
        e_digits, e_code = extract_digits(e_inv), clean_invoice_code(e_inv)

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue
            v_inv, v_amt = str(v.get("invoice_ven", "")).strip(), round(float(v["__amt"]), 2)
            v_digits, v_code = extract_digits(v_inv), clean_invoice_code(v_inv)
            diff, amt_close = round(e_amt - v_amt, 2), abs(round(e_amt - v_amt, 2)) < 0.05

            same_full = (e_inv.replace(" ", "") == v_inv.replace(" ", ""))
            same_clean = (e_code == v_code)
            suffix_ok = (e_code.endswith(v_code) or v_code.endswith(e_code))
            numeric_tail_ok = (e_digits.endswith(v_digits) or v_digits.endswith(e_digits))
            same_type = (e["__doctype"] == v["__doctype"])

            if same_type and (same_full or same_clean or suffix_ok or numeric_tail_ok) and amt_close:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match" if amt_close else "Difference"
                })
                used_vendor_rows.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = {m["ERP Invoice"] for _, m in matched_df.iterrows()}
    matched_ven = {m["Vendor Invoice"] for _, m in matched_df.iterrows()}

    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]] \
        if "invoice_ven" in ven_use else pd.DataFrame()
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]] \
        if "invoice_erp" in erp_use else pd.DataFrame()
    missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})

    return matched_df, missing_in_erp, missing_in_vendor, canceled_erp_df, canceled_ven_df


# ---------- PAYMENT DETECTION ----------
def extract_payments(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    payment_keywords = ["πληρωμή", "payment", "bank transfer", "transferencia", "pago", "remesa", "trf"]
    exclude_keywords = ["τιμολόγιο", "invoice", "παραστατικό", "expenses", "διόρθωση"]

    def is_payment(reason: str):
        text = str(reason or "").lower()
        return any(k in text for k in payment_keywords) and not any(x in text for x in exclude_keywords)

    erp_pay = erp_df[erp_df["reason_erp"].apply(is_payment)].copy() if "reason_erp" in erp_df else pd.DataFrame()
    ven_pay = ven_df[ven_df["reason_ven"].apply(is_payment)].copy() if "reason_ven" in ven_df else pd.DataFrame()

    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_erp")) - normalize_number(r.get("credit_erp"))), axis=1)
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_ven")) - normalize_number(r.get("credit_ven"))), axis=1)

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
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)

    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing, canceled_erp, canceled_ven = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)

    st.success("✅ Reconciliation complete")

    # ---------- Display ----------
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)

    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    st.dataframe(erp_missing, use_container_width=True)

    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    st.dataframe(ven_missing, use_container_width=True)

    st.subheader("🗑 Fully Canceled Invoices")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**ERP Canceled Pairs**")
        st.dataframe(canceled_erp, use_container_width=True)
    with col2:
        st.markdown("**Vendor Canceled Pairs**")
        st.dataframe(canceled_ven, use_container_width=True)

    # ---------- PAYMENTS ----------
    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**💼 ERP Payments**")
        st.dataframe(erp_pay, use_container_width=True)
        st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR" if not erp_pay.empty else "No ERP payments found.")
    with col2:
        st.markdown("**🧾 Vendor Payments**")
        st.dataframe(ven_pay, use_container_width=True)
        st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR" if not ven_pay.empty else "No Vendor payments found.")

    st.markdown("### ✅ Matched Payments")
    st.dataframe(matched_pay, use_container_width=True)
    st.markdown(f"**Total Matched ERP Payments:** {matched_pay['ERP Amount'].sum():,.2f} EUR")
    st.markdown(f"**Total Matched Vendor Payments:** {matched_pay['Vendor Amount'].sum():,.2f} EUR")
