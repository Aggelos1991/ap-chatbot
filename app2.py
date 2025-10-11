import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="🌍 Universal Vendor Reconciliation", layout="wide")
st.title("🌍 Universal Vendor Reconciliation App")

# ==========================================
# UNIVERSAL COLUMN DETECTION
# ==========================================
COLS = {
    "vendor": ["Vendor", "Supplier", "Supplier Name", "Προμηθευτής"],
    "trn": ["TRN", "AFM", "ΑΦΜ", "VAT", "CIF", "Tax ID"],
    "invoice": [
        "Invoice", "Invoice No", "Inv No", "Alt Document",
        "Alternative Document", "Αρ. Τιμολογίου", "Παραστατικό"
    ],
    "amount": ["Amount", "Value", "Invoice Value", "Ποσό"],
    "balance": ["Balance", "Υπόλοιπο", "Saldo"],
    "date": ["Date", "Ημερομηνία", "Fecha"]
}

def find_col(df, aliases):
    for col in df.columns:
        if str(col).strip().lower() in [a.lower() for a in aliases]:
            return col
    return None

# ==========================================
# SMART INVOICE NORMALIZATION & MATCHING
# ==========================================
def normalize_invoice(inv):
    """Extracts the numeric core of an invoice for fuzzy comparison."""
    s = str(inv).strip().upper()
    digits = re.sub(r"\D", "", s)
    return digits or s[-5:]  # fallback to last few chars if no digits

def invoice_match(erp_inv, ven_inv):
    """Smart comparison: matches if numbers overlap or end with same digits."""
    e = normalize_invoice(erp_inv)
    v = normalize_invoice(ven_inv)
    if not e or not v:
        return False
    return e == v or e.endswith(v) or v.endswith(e)

# ==========================================
# LOAD FILE
# ==========================================
def load_excel(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file)
        return df
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        return pd.DataFrame()

# ==========================================
# MATCHING LOGIC
# ==========================================
def match_invoices(erp_df, ven_df):
    matched_rows = []
    erp_unmatched, ven_unmatched = [], []

    # Detect universal columns in both dataframes
    cols = {}
    for key in COLS:
        cols[key + "_erp"] = find_col(erp_df, COLS[key])
        cols[key + "_ven"] = find_col(ven_df, COLS[key])

    missing_cols = [k for k, v in cols.items() if v is None and not k.startswith("amount")]
    if missing_cols:
        st.warning(f"⚠️ Missing columns detected: {missing_cols}. Matching might be incomplete.")

    for _, erp_row in erp_df.iterrows():
        erp_trn = str(erp_row.get(cols["trn_erp"], "")).strip()
        erp_inv = erp_row.get(cols["invoice_erp"], "")
        erp_vendor = erp_row.get(cols["vendor_erp"], "Unknown")
        erp_balance = float(erp_row.get(cols["balance_erp"], 0))

        # Find vendor matches by TRN/AFM/VAT
        candidates = ven_df[
            ven_df[cols["trn_ven"]].astype(str).str.strip() == erp_trn
        ] if cols["trn_ven"] else pd.DataFrame()

        found = False

        for _, ven_row in candidates.iterrows():
            ven_inv = ven_row.get(cols["invoice_ven"], "")
            if invoice_match(erp_inv, ven_inv):
                ven_balance = float(ven_row.get(cols["balance_ven"], 0))
                diff = round(erp_balance - ven_balance, 2)
                status = "✅ Match" if abs(diff) < 0.05 else "⚠️ Balance Difference"
                matched_rows.append({
                    "Vendor/Supplier": erp_vendor,
                    "TRN/AFM": erp_trn,
                    "ERP Invoice": erp_inv,
                    "Vendor Invoice": ven_inv,
                    "ERP Balance": erp_balance,
                    "Vendor Balance": ven_balance,
                    "Difference": diff,
                    "Status": status
                })
                found = True
                break

        if not found:
            erp_unmatched.append(erp_row)

    # vendor unmatched
    for _, ven_row in ven_df.iterrows():
        trn = str(ven_row.get(cols["trn_ven"], "")).strip()
        inv = ven_row.get(cols["invoice_ven"], "")
        in_erp = any(
            invoice_match(x, inv) and (str(y).strip() == trn)
            for x, y in zip(erp_df[cols["invoice_erp"]], erp_df[cols["trn_erp"]])
        )
        if not in_erp:
            ven_unmatched.append(ven_row)

    return pd.DataFrame(matched_rows), pd.DataFrame(erp_unmatched), pd.DataFrame(ven_unmatched)

# ==========================================
# STREAMLIT UI
# ==========================================
st.write("Upload your ERP export and Vendor statement below (any language or column names supported):")

erp_file = st.file_uploader("📘 Upload ERP Export (Excel)", type=["xlsx"])
vendor_file = st.file_uploader("📗 Upload Vendor Statement (Excel)", type=["xlsx"])

if erp_file and vendor_file:
    erp_df = load_excel(erp_file)
    ven_df = load_excel(vendor_file)
    with st.spinner("Reconciling..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    st.success("✅ Reconciliation complete!")

    st.subheader("📊 Matched / Differences")
    st.dataframe(matched)

    st.subheader("❌ In ERP but Missing in Vendor")
    st.dataframe(erp_missing)

    st.subheader("❌ In Vendor but Missing in ERP")
    st.dataframe(ven_missing)

    # Downloads
    st.download_button("⬇️ Download Matched", matched.to_csv(index=False).encode(), "matched.csv")
    st.download_button("⬇️ Download Missing in ERP", ven_missing.to_csv(index=False).encode(), "missing_in_erp.csv")
    st.download_button("⬇️ Download Missing in Vendor", erp_missing.to_csv(index=False).encode(), "missing_in_vendor.csv")
