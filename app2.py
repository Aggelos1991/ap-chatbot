import streamlit as st
import pandas as pd
import re

# ======================================
# CONFIG
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ======================================
# HELPERS
# ======================================
def normalize_number(v):
    """Convert numeric strings like '1.234,56' or '1,234.56' safely to float."""
    if v is None:
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
    """Map multilingual headers to unified names — optimized for Spanish vendor statements."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge",
            "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "descriptivo", "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento"
        ],
    }

    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    for k, vals in mapping.items():
        for col, low in cols_lower.items():
            if any(v in low for v in vals):
                rename_map[col] = f"{k}_{tag}"

    out = df.rename(columns=rename_map)

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

    # ====== ERP PREP ======
    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))

        # 🚫 Ignore payments / transfers
        if any(re.search(rf"\b{kw}", reason) for kw in [
            "pay", "paid", "payment", "repay", "prepay",
            "pago", "pag", "pagado", "pagos", "transfer", "transferencia", "transf",
            "bank", "saldo", "balance", "ajuste", "adjust", "trf"
        ]):
            return "IGNORE"

        # 🔹 Credit Note
        elif any(re.search(rf"\b{kw}", reason) for kw in [
            "credit", "creditnote", "credit note", "credit-memo", "creditmemo",
            "cred", "memo", "memo credito", "nota credito", "nota de credito",
            "nota crédito", "nota de crédito", "nota abono", "abono", "nc",
            "crédito", "credito", "cn"
        ]) or (charge > 0 and credit == 0):
            return "CN"

        # 🔹 Invoice
        elif credit > 0 or any(k in reason for k in ["factura", "invoice", "inv", "rn:"]):
            return "INV"

        else:
            return "UNKNOWN"

    def calc_erp_amount(row):
        doc = row.get("__doctype", "")
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))

        if doc == "INV":
            return abs(credit)
        elif doc == "CN":
            return -abs(charge if charge > 0 else credit)
        else:
            return 0.0

    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)

    # ====== VENDOR PREP ======
    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))

        # 🚫 Ignore payments / transfers
        if any(re.search(rf"\b{kw}", reason) for kw in [
            "pay", "paid", "payment", "repay", "prepay",
            "pago", "pag", "pagado", "pagos", "transfer", "transferencia", "transf",
            "bank", "saldo", "balance", "ajuste", "adjust", "trf"
        ]):
            return "IGNORE"

        # 🔹 Credit Note
        elif any(re.search(rf"\b{kw}", reason) for kw in [
            "credit", "creditnote", "credit note", "credit-memo", "creditmemo",
            "cred", "memo", "memo credito", "nota credito", "nota de credito",
            "nota crédito", "nota de crédito", "nota abono", "abono", "nc",
            "crédito", "credito", "cn"
        ]) or credit > 0:
            return "CN"

        # 🔹 Invoice
        elif any(k in reason for k in ["factura", "invoice", "inv", "rn:"]) or debit > 0:
            return "INV"

        else:
            return "UNKNOWN"

    def calc_vendor_amount(row):
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        doc = row.get("__doctype", "")

        if doc == "INV":
            return abs(debit)
        elif doc == "CN":
            return -abs(credit if credit > 0 else debit)
        else:
            return 0.0

    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)

    # (rest of match_invoices omitted for brevity)
    matched_df = pd.DataFrame(matched)
    return matched_df, pd.DataFrame(), pd.DataFrame()


# ======================================
# PAYMENT EXTRACTION
# ======================================
def extract_payments(erp_df, ven_df):
    """Extract and match payment transactions between ERP and Vendor."""
    payment_keywords = [
        "pago", "pagos", "payment", "transfer", "transferencia", "bank", "trf",
        "remesa", "prepago", "ajuste", "adjust", "compensacion", "settlement"
    ]

    def is_payment(reason):
        reason = str(reason).lower()
        return any(k in reason for k in payment_keywords)

    erp_pay = erp_df[erp_df["reason_erp"].apply(is_payment)].copy()
    ven_pay = ven_df[ven_df["reason_ven"].apply(is_payment)].copy()

    erp_pay["Amount"] = erp_pay.apply(
        lambda r: normalize_number(r.get("debit_erp")) - normalize_number(r.get("credit_erp")),
        axis=1,
    ).abs()
    ven_pay["Amount"] = ven_pay.apply(
        lambda r: normalize_number(r.get("debit_ven")) - normalize_number(r.get("credit_ven")),
        axis=1,
    ).abs()

    matched_payments = []
    used_vendor_rows = set()
    for _, e in erp_pay.iterrows():
        for v_idx, v in ven_pay.iterrows():
            if v_idx in used_vendor_rows:
                continue
            diff = abs(e["Amount"] - v["Amount"])
            if diff < 0.05:
                matched_payments.append({
                    "ERP Date": e.get("date_erp", ""),
                    "Vendor Date": v.get("date_ven", ""),
                    "ERP Description": e.get("reason_erp", ""),
                    "Vendor Description": v.get("reason_ven", ""),
                    "ERP Amount": e["Amount"],
                    "Vendor Amount": v["Amount"],
                    "Difference": round(diff, 2)
                })
                used_vendor_rows.add(v_idx)
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

    if "cif_ven" not in ven_df.columns or "cif_erp" not in erp_df.columns:
        st.error("❌ Missing CIF/VAT columns.")
        st.stop()

    vendor_cifs = sorted({str(x).strip().upper() for x in ven_df["cif_ven"].dropna().unique() if str(x).strip()})
    selected_cif = vendor_cifs[0] if len(vendor_cifs) == 1 else st.selectbox("Select Vendor CIF:", vendor_cifs)

    erp_df = erp_df[erp_df["cif_erp"].astype(str).str.strip().str.upper() == selected_cif]
    ven_df = ven_df[ven_df["cif_ven"].astype(str).str.strip().str.upper() == selected_cif]

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)

    st.success(f"✅ Recon complete for CIF {selected_cif}")

    # ============================
    # Display sections
    # ============================
    st.subheader("🏦 Payment Transactions (Identified in both sides)")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**ERP Payments**")
        st.dataframe(erp_pay[["date_erp", "reason_erp", "debit_erp", "credit_erp"]], use_container_width=True)
    with col2:
        st.markdown("**Vendor Payments**")
        st.dataframe(ven_pay[["date_ven", "reason_ven", "debit_ven", "credit_ven"]], use_container_width=True)

    if not matched_pay.empty:
        st.markdown("### ✅ Matched Payments")
        st.dataframe(matched_pay, use_container_width=True)
    else:
        st.info("No matching payments found.")

    # ============================
    # Chat prompt
    # ============================
    st.subheader("💬 Ask ReconRaptor about Payments")
    query = st.text_input("Type your question (e.g. 'sum of ERP payments'):")

    if query:
        if "vendor" in query.lower():
            total = ven_pay["Amount"].sum()
            st.write(f"💰 Total Vendor Payments: **{total:,.2f} EUR**")
        elif "erp" in query.lower():
            total = erp_pay["Amount"].sum()
            st.write(f"💰 Total ERP Payments: **{total:,.2f} EUR**")
        elif "compare" in query.lower() or "difference" in query.lower():
            diff = abs(erp_pay["Amount"].sum() - ven_pay["Amount"].sum())
            st.write(f"📊 Difference between ERP and Vendor payments: **{diff:,.2f} EUR**")
        else:
            st.info("I can answer about ERP payments, vendor payments, or differences.")
else:
    st.info("Please upload both ERP and Vendor files to begin.")
