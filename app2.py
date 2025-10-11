import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz

st.set_page_config(page_title="🤝 Vendor Reconciliation", layout="wide")
st.title("🧾 Universal Vendor Reconciliation App")

# ============================================================
# Universal Column Mapping — multilingual support (EN/ES/GR)
# ============================================================
def normalize_columns(df, source="ven"):
    colmap = {
        "vendor": ["supplier name", "vendor", "proveedor", "προμηθευτής"],
        "trn": ["tax id", "cif", "vat", "afm", "trn", "vat number", "tax number"],
        "invoice": ["invoice number", "alt document", "invoice", "παραστατικό", "factura"],
        "balance": ["balance", "saldo", "υπόλοιπο", "amount", "importe", "valor"],
        "date": ["date", "fecha", "ημερομηνία"],
    }

    rename_map = {}
    for key, variants in colmap.items():
        for col in df.columns:
            col_low = str(col).strip().lower()
            if any(v in col_low for v in variants):
                rename_map[col] = f"{key}_{source}"
                break

    df = df.rename(columns=rename_map)
    return df


# ============================================================
# Matching Logic
# ============================================================
def match_invoices(erp_df, ven_df):
    matched = []
    erp_missing = []
    ven_missing = ven_df.copy()

    for _, erp_row in erp_df.iterrows():
        erp_trn = str(erp_row.get("trn_erp", "")).strip()
        erp_invoice = str(erp_row.get("invoice_erp", "")).strip()
        erp_balance = float(erp_row.get("balance_erp", 0))

        # Find vendor in vendor dataframe
        ven_subset = ven_df[ven_df["trn_ven"] == erp_trn]
        found = False
        for _, ven_row in ven_subset.iterrows():
            ven_invoice = str(ven_row.get("invoice_ven", "")).strip()
            ven_balance = float(ven_row.get("balance_ven", 0))

            # Flexible matching (last digits, partials, fuzzy)
            if (
                erp_invoice[-4:] in ven_invoice
                or ven_invoice[-4:] in erp_invoice
                or fuzz.ratio(erp_invoice, ven_invoice) > 75
            ):
                diff = round(erp_balance - ven_balance, 2)
                status = "Match" if diff == 0 else "Difference"
                matched.append({
                    "Vendor/Supplier": erp_row.get("vendor_erp", ""),
                    "TRN/AFM": erp_trn,
                    "ERP Invoice": erp_invoice,
                    "Vendor Invoice": ven_invoice,
                    "ERP Balance": erp_balance,
                    "Vendor Balance": ven_balance,
                    "Difference": diff,
                    "Status": status
                })
                ven_missing = ven_missing.drop(ven_row.name)
                found = True
                break

        if not found:
            erp_missing.append(erp_row)

    return pd.DataFrame(matched), pd.DataFrame(erp_missing), ven_missing


# ============================================================
# Streamlit interface
# ============================================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = pd.read_excel(uploaded_erp)
    ven_df = pd.read_excel(uploaded_vendor)

    # Normalize column names
    erp_df = normalize_columns(erp_df, source="erp")
    ven_df = normalize_columns(ven_df, source="ven")

    # Check if required columns exist
    required_cols = ["trn_erp", "invoice_erp", "balance_erp", "vendor_erp", "trn_ven", "invoice_ven", "balance_ven"]
    missing_cols = [c for c in required_cols if c not in erp_df.columns and c not in ven_df.columns]

    if missing_cols:
        st.warning(f"⚠️ Missing columns detected: {missing_cols}. Matching might be incomplete.")

    with st.spinner("🔍 Reconciling... please wait..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    st.success("✅ Reconciliation complete!")

    st.subheader("📊 Matched / Differences")
    st.dataframe(matched)

    st.subheader("❌ In ERP but Missing in Vendor")
    st.dataframe(erp_missing)

    st.subheader("❌ In Vendor but Missing in ERP")
    st.dataframe(ven_missing)

    st.download_button("⬇️ Download Matched", matched.to_csv(index=False).encode("utf-8"), "matched.csv", "text/csv")

else:
    st.info("Please upload both ERP Export and Vendor Statement files to begin.")
