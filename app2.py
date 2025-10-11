import streamlit as st
import pandas as pd
import re

# ==========================
# ReconRaptor setup 🦖
# ==========================
st.set_page_config(page_title="🦖 ReconRaptor", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ==========================
# Helper functions
# ==========================
def normalize_number(v):
    """Convert formatted numbers to float safely."""
    if pd.isna(v):
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
    """Unify multilingual headers for ERP or Vendor."""
    mapping = {
        "vendor": ["supplier", "vendor", "proveedor", "προμηθευτής"],
        "trn": ["vat", "cif", "afm", "trn", "tax id"],
        "invoice": ["invoice", "alt document", "factura", "παραστατικό"],
        "description": ["description", "descripción", "περιγραφή"],
        "debit": ["debit", "debe", "χρέωση"],
        "credit": ["credit", "haber", "πίστωση"],
        "amount": ["amount", "importe", "valor"],
        "balance": ["balance", "saldo", "υπόλοιπο"],
        "date": ["date", "fecha", "ημερομηνία"],
    }

    rename_map = {}
    for key, synonyms in mapping.items():
        for col in df.columns:
            c = str(col).strip().lower()
            if any(word in c for word in synonyms):
                rename_map[col] = f"{key}_{tag}"
    return df.rename(columns=rename_map)


def normalize_invoice_num(s):
    """Standardize invoice numbers for comparison."""
    return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()


# ==========================
# Core reconciliation
# ==========================
def match_invoices(erp_df, ven_df):
    matched_rows = []
    matched_erp_invoices = set()
    matched_ven_invoices = set()

    # ---- detect key columns ----
    trn_erp = next((c for c in erp_df.columns if "trn_" in c), None)
    trn_ven = next((c for c in ven_df.columns if "trn_" in c), None)
    inv_erp = next((c for c in erp_df.columns if "invoice_" in c), None)
    inv_ven = next((c for c in ven_df.columns if "invoice_" in c), None)
    desc_ven = next((c for c in ven_df.columns if "description_" in c), None)

    # safety guard
    if not inv_erp or not inv_ven:
        return pd.DataFrame(), erp_df, ven_df

    # ---- normalize invoice numbers ----
    erp_df["__inv"] = erp_df[inv_erp].astype(str).map(normalize_invoice_num)
    ven_df["__inv"] = ven_df[inv_ven].astype(str).map(normalize_invoice_num)

    # restrict to same TRN/vendor if available
    if trn_erp and trn_ven and not ven_df[trn_ven].dropna().empty:
        vendor_trn = str(ven_df[trn_ven].dropna().iloc[0])
        erp_df = erp_df[erp_df[trn_erp].astype(str) == vendor_trn]

    # remove payments from vendor
    pay_words = ["pago", "payment", "transfer", "partial", "liquidación"]
    if desc_ven and desc_ven in ven_df.columns:
        ven_df = ven_df[~ven_df[desc_ven].astype(str).str.lower().apply(lambda x: any(w in x for w in pay_words))]

    # ---- matching loop ----
    for _, e_row in erp_df.iterrows():
        e_inv = e_row["__inv"]
        e_amt = normalize_number(e_row.get("amount_erp", e_row.get("debit_erp", e_row.get("credit_erp", 0))))

        # match vendor invoice by number
        match = ven_df[ven_df["__inv"] == e_inv]
        if not match.empty:
            v_row = match.iloc[0]
            desc = str(v_row.get(desc_ven, "")).lower()
            v_debit = normalize_number(v_row.get("debit_ven", 0))
            v_credit = normalize_number(v_row.get("credit_ven", 0))
            v_amt = v_credit if "abono" in desc or "credit" in desc else v_debit

            # ✅ Fix: handle credit notes with opposite signs
            if (e_amt < 0 and v_amt > 0) or (e_amt > 0 and v_amt < 0):
                v_amt = -v_amt  # flip vendor amount to align with ERP

            diff = round(e_amt - v_amt, 2)
            status = "Match" if abs(diff) < 0.05 else "Difference"

            matched_rows.append({
                "Vendor/Supplier": e_row.get("vendor_erp", ""),
                "TRN/AFM": e_row.get("trn_erp", ""),
                "ERP Invoice": e_row.get(inv_erp, ""),
                "Vendor Invoice": v_row.get(inv_ven, ""),
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status,
                "Description": v_row.get(desc_ven, "")
            })
            matched_erp_invoices.add(e_inv)
            matched_ven_invoices.add(e_inv)

    matched = pd.DataFrame(matched_rows)

    # ✅ Correct missing logic
    # → If it's in vendor but not ERP → Missing in ERP
    # → If it's in ERP but not vendor → Missing in Vendor
    erp_missing = ven_df[~ven_df["__inv"].isin(matched_erp_invoices)].reset_index(drop=True)
    ven_missing = erp_df[~erp_df["__inv"].isin(matched_ven_invoices)].reset_index(drop=True)

    return matched, erp_missing, ven_missing


# ==========================
# Streamlit UI 🦖
# ==========================
st.write("Upload your ERP Export and Vendor Statement for reconciliation:")

uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(pd.read_excel(uploaded_erp), "erp")
    ven_df = normalize_columns(pd.read_excel(uploaded_vendor), "ven")

    with st.spinner("🦖 ReconRaptor is reconciling your invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    total_match = len(matched[matched["Status"] == "Match"]) if not matched.empty else 0
    total_diff = len(matched[matched["Status"] == "Difference"]) if not matched.empty else 0
    total_missing = len(erp_missing) + len(ven_missing)
    st.success(f"✅ Recon complete: {total_match} matched, {total_diff} differences, {total_missing} missing")

    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        else:
            return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    st.dataframe(matched.style.apply(highlight_row, axis=1))

    st.subheader("❌ Missing in ERP")
    st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"))

    st.subheader("❌ Missing in Vendor")
    st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"))

    st.download_button(
        "⬇️ Download Matched CSV",
        matched.to_csv(index=False).encode("utf-8"),
        "ReconRaptor_Results.csv",
        "text/csv"
    )
else:
    st.info("🦖 Please upload both ERP and Vendor Statement files to begin.")
