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
    """Map multilingual headers to unified names (English + Spanish + Greek)."""
    mapping = {
        "invoice": [
            "invoice", "factura", "τιμολόγιο", "παραστατικό", "document", "doc", "ref",
            "referencia", "nº", "num", "numero", "número", "nº factura", "num factura", "alternative document"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "πιστωτικό", "πίστωση", "απόδειξη επιστροφής", "importe haber", "valor haber"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "ποσό", "ποσότητα", "σύνολο",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción", "αιτία", "περιγραφή",
            "detalle", "detalles", "razon", "razón", "παρατήρηση", "σχόλιο", "σχόλια", "explicacion"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "afm", "αφμ", "tax", "id fiscal", "número fiscal", "num fiscal", "code"
        ],
        "date": [
            "date", "fecha", "ημερομηνία", "data", "fecha factura", "fecha documento", "fecha doc"
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
# AGGREGATION (CANCELLATIONS, REPOSTINGS, RETURNS)
# ======================================
def aggregate_invoice_amounts(df):
    """Aggregates invoices (Alternative Document + CIF) and computes the net remaining amount."""
    alt_doc_col = next((c for c in df.columns if "invoice" in c.lower() or "alternative document" in c.lower()), None)
    cif_col = next((c for c in df.columns if "cif" in c.lower() or "nif" in c.lower() or "vat" in c.lower() or "afm" in c.lower()), None)
    charge_col = next((c for c in df.columns if "debit" in c.lower() or "charge" in c.lower()), None)
    credit_col = next((c for c in df.columns if "credit" in c.lower()), None)

    if not all([alt_doc_col, cif_col, charge_col, credit_col]):
        st.warning("⚠️ Missing columns for aggregation — skipping partial return logic.")
        return df

    df[charge_col] = pd.to_numeric(df[charge_col], errors="coerce").fillna(0.0)
    df[credit_col] = pd.to_numeric(df[credit_col], errors="coerce").fillna(0.0)

    grouped = (
        df.groupby([alt_doc_col, cif_col], as_index=False)
        .agg({charge_col: "sum", credit_col: "sum"})
    )

    grouped["Net_Amount"] = grouped[charge_col] - grouped[credit_col]
    grouped = grouped[abs(grouped["Net_Amount"]) > 0.01]
    grouped["Status"] = grouped["Net_Amount"].apply(lambda x: "Refund" if x < 0 else "Outstanding")

    df = df.merge(
        grouped[[alt_doc_col, cif_col, "Net_Amount", "Status"]],
        on=[alt_doc_col, cif_col],
        how="left"
    )

    return df


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    # Multilingual keywords
    PAYMENT_WORDS = ["pago", "payment", "transfer", "bank", "saldo", "trf", "πληρωμή", "μεταφορά", "τράπεζα", "υπόλοιπο"]
    CREDIT_WORDS = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση"]
    INVOICE_WORDS = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]

    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if any(k in reason for k in PAYMENT_WORDS):
            return "IGNORE"
        elif any(k in reason for k in CREDIT_WORDS):
            return "CN"
        elif any(k in reason for k in INVOICE_WORDS) or credit > 0:
            return "INV"
        return "UNKNOWN"

    def calc_erp_amount(row):
        if "Net_Amount" in row and not pd.isna(row["Net_Amount"]):
            return float(row["Net_Amount"])
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
        if any(k in reason for k in PAYMENT_WORDS):
            return "IGNORE"
        elif any(k in reason for k in CREDIT_WORDS) or credit > 0:
            return "CN"
        elif any(k in reason for k in INVOICE_WORDS) or debit > 0:
            return "INV"
        return "UNKNOWN"

    def calc_vendor_amount(row):
        if "Net_Amount" in row and not pd.isna(row["Net_Amount"]):
            return float(row["Net_Amount"])
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        doc = row.get("__doctype", "")
        if doc == "INV":
            return abs(debit)
        elif doc == "CN":
            return -abs(credit if credit > 0 else debit)
        return 0.0

    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    def extract_digits(v):
        return re.sub(r"\D", "", str(v or "")).lstrip("0")

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_digits = extract_digits(e_inv)
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_digits = extract_digits(v_inv)
            diff = round(e_amt - v_amt, 2)
            amt_close = abs(diff) < 0.05
            if e_inv == v_inv or (e_digits and v_digits and (e_digits.endswith(v_digits) or v_digits.endswith(e_digits))):
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

    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]] if "invoice_ven" in ven_use else pd.DataFrame()
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]] if "invoice_erp" in erp_use else pd.DataFrame()

    missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})

    return matched_df, missing_in_erp, missing_in_vendor


# ======================================
# PAYMENT EXTRACTION (EN + ES + GR)
# ======================================
def extract_payments(erp_df, ven_df):
    PAYMENT_KEYWORDS = [
        "pago", "pagos", "payment", "transfer", "transferencia", "bank", "trf", "remesa", "prepago", "ajuste",
        "πληρωμή", "μεταφορά", "τραπεζα", "έμβασμα", "προκαταβολή", "επιστροφή"
    ]
    is_payment = lambda x: any(k in str(x).lower() for k in PAYMENT_KEYWORDS)

    erp_pay = erp_df[erp_df["reason_erp"].apply(is_payment)].copy() if "reason_erp" in erp_df else pd.DataFrame()
    ven_pay = ven_df[ven_df["reason_ven"].apply(is_payment)].copy() if "reason_ven" in ven_df else pd.DataFrame()

    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(lambda r: abs(normalize_number(r.get("debit_erp")) - normalize_number(r.get("credit_erp"))), axis=1)
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(lambda r: abs(normalize_number(r.get("debit_ven")) - normalize_number(r.get("credit_ven"))), axis=1)

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
                    "ERP Amount": e["Amount"],
                    "Vendor Amount": v["Amount"],
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

    # ✅ Handle cancellations / repostings / partial returns
    erp_df = aggregate_invoice_amounts(erp_df)
    ven_df = aggregate_invoice_amounts(ven_df)

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)

    st.success("✅ Reconciliation complete")

    # ====== HIGHLIGHTING ======
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        return [''] * len(row)

    # ====== MATCHED ======
    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")

    # ====== MISSING ======
    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in Vendor.")

    # ====== PAYMENTS ======
    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**💼 ERP Payments**")
        if not erp_pay.empty:
            st.dataframe(erp_pay.style.applymap(lambda _: "background-color: #004d40; color: white"), use_container_width=True)
            st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No ERP payments found.")
    with col2:
        st.markdown("**🧾 Vendor Payments**")
        if not ven_pay.empty:
            st.dataframe(ven_pay.style.applymap(lambda _: "background-color: #1565c0; color: white"), use_container_width=True)
            st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No Vendor payments found.")

    st.markdown("### ✅ Matched Payments")
    if not matched_pay.empty:
        st.dataframe(matched_pay.style.applymap(lambda _: "background-color: #2e7d32; color: white"), use_container_width=True)
        total_erp = matched_pay["ERP Amount"].sum()
        total_vendor = matched_pay["Vendor Amount"].sum()
        diff_total = abs(total_erp - total_vendor)
        st.markdown(f"**Total Matched ERP Payments:** {total_erp:,.2f} EUR")
        st.markdown(f"**Total Matched Vendor Payments:** {total_vendor:,.2f} EUR")
        st.markdown(f"**Difference Between ERP and Vendor Payments:** {diff_total:,.2f} EUR")
    else:
        st.info("No matching payments found.")
