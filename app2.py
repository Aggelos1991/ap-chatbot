import streamlit as st
import pandas as pd
import re
from datetime import datetime

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
            "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
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

    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if any(k in reason for k in ["pago", "payment", "transfer", "bank", "saldo", "trf"]):
            return "IGNORE"
        elif any(k in reason for k in ["credit", "nota", "abono", "cn"]):
            return "CN"
        elif credit > 0 or any(k in reason for k in ["factura", "invoice", "inv"]):
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
        if any(k in reason for k in ["pago", "payment", "transfer", "bank", "saldo", "trf"]):
            return "IGNORE"
        elif any(k in reason for k in ["credit", "nota", "abono", "cn"]) or credit > 0:
            return "CN"
        elif any(k in reason for k in ["factura", "invoice", "inv"]) or debit > 0:
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
# PAYMENT EXTRACTION
# ======================================
def extract_payments(erp_df, ven_df):
    keywords = ["pago", "pagos", "payment", "transfer", "transferencia", "bank", "trf", "remesa", "prepago", "ajuste"]
    is_payment = lambda x: any(k in str(x).lower() for k in keywords)

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
    # ====== PREVIOUS YEAR CHECK ======
    if "reason_erp" in erp_df.columns:
        erp_df["reason_erp"] = erp_df["reason_erp"].astype(str).str.strip().str.lower()
        if "credit_erp" in erp_df.columns or "debit_erp" in erp_df.columns:
            credit_vals = erp_df["credit_erp"].apply(normalize_number) if "credit_erp" in erp_df else pd.Series(0)
            debit_vals = erp_df["debit_erp"].apply(normalize_number) if "debit_erp" in erp_df else pd.Series(0)
    
            prev_mask = erp_df["reason_erp"].str.contains("previous year", na=False)
            combined_vals = credit_vals.combine_first(debit_vals)
    
            # Filter only the rows with "previous year"
            prev_rows = combined_vals[prev_mask]
    
            if not prev_rows.empty:
                total_prev = prev_rows.sum()
                if total_prev > 0:
                    st.warning(f"⚠️ You have open amounts from previous years totaling **{total_prev:,.2f} EUR**.")
                    if not st.button("✅ OK, continue"):
                        st.stop()
                elif total_prev < 0:
                    st.warning(f"⚠️ You have a balance carried from previous years totaling **{total_prev:,.2f} EUR**.")
                    if not st.button("✅ OK, continue"):
                        st.stop()



    if "cif_erp" not in erp_df.columns or "cif_ven" not in ven_df.columns:
        st.error("❌ Missing CIF/VAT columns.")
        st.stop()

    vendor_cifs = sorted({str(x).strip().upper() for x in ven_df["cif_ven"].dropna().unique()})
    selected_cif = vendor_cifs[0] if len(vendor_cifs) == 1 else st.selectbox("Select Vendor CIF:", vendor_cifs)

    erp_df = erp_df[erp_df["cif_erp"].astype(str).str.upper() == selected_cif]
    ven_df = ven_df[ven_df["cif_ven"].astype(str).str.upper() == selected_cif]

    # ====== PREVIOUS YEAR CHECK ======
    if "reason_erp" in erp_df.columns and "credit_erp" in erp_df.columns:
        prev_mask = erp_df["reason_erp"].astype(str).str.lower().str.contains("previous year")
        credit_vals = erp_df["credit_erp"].apply(normalize_number)
        if any(prev_mask & (credit_vals > 0)):
            st.warning("⚠️ You have open amounts from previous years. Click OK to continue.")
            if not st.button("✅ OK, continue"):
                st.stop()

    # ====== RECONCILIATION ======
    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)

    st.success(f"✅ Reconciliation complete for CIF {selected_cif}")

    # ====== HIGHLIGHTING ======
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        return [''] * len(row)

    # ====== TABLES ======
    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")

    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(
            erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True,
        )
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(
            ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True,
        )
    else:
        st.success("✅ No missing invoices in Vendor.")

    # ====== PAYMENTS ======
    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**ERP Payments**")
        st.dataframe(erp_pay, use_container_width=True)
    with col2:
        st.markdown("**Vendor Payments**")
        st.dataframe(ven_pay, use_container_width=True)

    if not matched_pay.empty:
        st.markdown("### ✅ Matched Payments")
        st.dataframe(matched_pay, use_container_width=True)
    else:
        st.info("No matching payments found.")

# ====== CHAT PROMPT ======
st.subheader("💬 Ask ReconRaptor about Payments")

query = st.text_input("Ask something (e.g. 'sum of ERP payments'):")

if query:
    if "vendor" in query.lower():
        total = ven_pay["Amount"].sum() if "Amount" in ven_pay else 0
        st.write(f"💰 Total Vendor Payments: **{total:,.2f} EUR**")

    elif "erp" in query.lower():
        total = erp_pay["Amount"].sum() if "Amount" in erp_pay else 0
        st.write(f"💰 Total ERP Payments: **{total:,.2f} EUR**")

    elif "difference" in query.lower() or "compare" in query.lower():
        diff = abs(
            (erp_pay["Amount"].sum() if "Amount" in erp_pay else 0) -
            (ven_pay["Amount"].sum() if "Amount" in ven_pay else 0)
        )
        st.write(f"📊 Difference between ERP and Vendor payments: **{diff:,.2f} EUR**")

    else:
        st.info("I can answer about ERP payments, vendor payments, or differences.")
else:
    st.info("Please upload both ERP and Vendor files to begin.")


