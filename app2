import streamlit as st
import pandas as pd
import re
from fuzzywuzzy import fuzz

st.set_page_config(page_title="🧾 Vendor Reconciliation App", layout="wide")
st.title("🧾 Vendor Reconciliation App")

def normalize_invoice(inv):
    """Keep last 3-4 digits only, remove non-numeric"""
    digits = re.sub(r"\D", "", str(inv))
    return digits[-4:] if len(digits) >= 3 else digits

def load_excel(uploaded_file):
    return pd.read_excel(uploaded_file)

def match_invoices(erp_df, ven_df):
    matched_rows = []
    erp_unmatched, ven_unmatched = [], []

    for _, erp_row in erp_df.iterrows():
        erp_trn = str(erp_row["TRN"]).strip()
        erp_inv = normalize_invoice(erp_row["Invoice"])
        erp_vendor = erp_row["Vendor"]
        erp_amount = float(erp_row["Amount"])

        # find potential vendor matches in vendor statement
        candidates = ven_df[ven_df["TRN"].astype(str).str.strip() == erp_trn]
        found = False

        for _, ven_row in candidates.iterrows():
            ven_inv = normalize_invoice(ven_row["Invoice"])
            if erp_inv == ven_inv:
                diff = round(float(erp_row["Balance"]) - float(ven_row["Balance"]), 2)
                status = "✅ Match" if abs(diff) < 0.05 else "⚠️ Balance Difference"
                matched_rows.append({
                    "Vendor": erp_vendor,
                    "TRN": erp_trn,
                    "ERP Invoice": erp_row["Invoice"],
                    "Vendor Invoice": ven_row["Invoice"],
                    "ERP Balance": erp_row["Balance"],
                    "Vendor Balance": ven_row["Balance"],
                    "Difference": diff,
                    "Status": status
                })
                found = True
                break

        if not found:
            erp_unmatched.append(erp_row)

    # vendor unmatched
    for _, ven_row in ven_df.iterrows():
        trn = str(ven_row["TRN"]).strip()
        inv = normalize_invoice(ven_row["Invoice"])
        in_erp = any(
            (normalize_invoice(x) == inv) and (str(y).strip() == trn)
            for x, y in zip(erp_df["Invoice"], erp_df["TRN"])
        )
        if not in_erp:
            ven_unmatched.append(ven_row)

    return pd.DataFrame(matched_rows), pd.DataFrame(erp_unmatched), pd.DataFrame(ven_unmatched)

st.write("Upload your ERP export and Vendor statement below:")

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
