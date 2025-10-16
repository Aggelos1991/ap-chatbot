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
def normalize_number(value):
    """Normalize decimals like 1.234,56 → 1234.56"""
    if not value:
        return 0.0
    s = str(value).strip().replace(" ", "")
    s = re.sub(r"[^\d,.,-]", "", s)
    if "," in s and "." in s:
        if s.find(",") > s.find("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0


# ======================================
# AGGREGATE & CLEAN INVOICE DUPLICATES
# ======================================
def aggregate_invoice_amounts(df):
    """
    Groups invoices (Alternative Document + CIF) and keeps only the net amount after cancellations or partial returns.
    Automatically ignores fully neutralized entries and keeps only the net effective invoice.
    """

    required = {"Alternative Document", "Charge", "Credit", "CIF"}
    if not required.issubset(df.columns):
        st.warning(f"⚠️ Missing columns: {required - set(df.columns)} — skipping aggregation.")
        return df

    # Convert numeric columns safely
    for col in ["Charge", "Credit"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Group by invoice & CIF
    grouped = (
        df.groupby(["Alternative Document", "CIF"], as_index=False)
        .agg({"Charge": "sum", "Credit": "sum"})
    )

    # Compute final net amount
    grouped["Net_Amount"] = grouped["Charge"] - grouped["Credit"]

    # Keep only invoices where something remains (not fully canceled)
    grouped = grouped[abs(grouped["Net_Amount"]) > 0.01]

    # Status label
    grouped["Status"] = grouped["Net_Amount"].apply(
        lambda x: "Refund" if x < 0 else "Outstanding"
    )

    # Merge back minimal info (to preserve CIF, Reason, etc.)
    df_clean = df.drop_duplicates(subset=["Alternative Document", "CIF"]).merge(
        grouped[["Alternative Document", "CIF", "Net_Amount", "Status"]],
        on=["Alternative Document", "CIF"],
        how="right",
    )

    return df_clean


# ======================================
# INVOICE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    """Basic matching logic — compares invoice numbers and amounts."""

    if erp_df.empty or ven_df.empty:
        st.error("❌ One of the files is empty.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Normalize invoice fields
    erp_df["invoice_erp"] = erp_df["Alternative Document"].astype(str).str.strip()
    ven_df["invoice_ven"] = ven_df["Alternative Document"].astype(str).str.strip()

    # Normalize amount field (use Net_Amount if available)
    erp_df["__amt"] = erp_df.get("Net_Amount", erp_df.get("Charge", 0) - erp_df.get("Credit", 0))
    ven_df["__amt"] = ven_df.get("Net_Amount", ven_df.get("Charge", 0) - ven_df.get("Credit", 0))

    # Round for safety
    erp_df["__amt"] = erp_df["__amt"].round(2)
    ven_df["__amt"] = ven_df["__amt"].round(2)

    # Match by exact invoice number and amount
    matched = pd.merge(
        erp_df,
        ven_df,
        left_on=["invoice_erp", "__amt"],
        right_on=["invoice_ven", "__amt"],
        how="inner",
        suffixes=("_erp", "_ven"),
    )

    # Build missing tables
    matched_invoices_erp = set(matched["invoice_erp"])
    matched_invoices_ven = set(matched["invoice_ven"])

    erp_missing = erp_df[~erp_df["invoice_erp"].isin(matched_invoices_erp)]
    ven_missing = ven_df[~ven_df["invoice_ven"].isin(matched_invoices_ven)]

    return matched, erp_missing, ven_missing


# ======================================
# FILE UPLOADS
# ======================================
st.subheader("📂 Upload ERP and Vendor Files")

erp_file = st.file_uploader("Upload ERP Excel file", type=["xlsx"])
ven_file = st.file_uploader("Upload Vendor Excel file", type=["xlsx"])

if erp_file and ven_file:
    erp_df = pd.read_excel(erp_file)
    ven_df = pd.read_excel(ven_file)

    # ✅ Clean up data before matching
    st.info("🧹 Cleaning and aggregating ERP data...")
    erp_df = aggregate_invoice_amounts(erp_df)
    st.info("🧹 Cleaning and aggregating Vendor data...")
    ven_df = aggregate_invoice_amounts(ven_df)

    # ✅ Perform matching
    matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    # ======================================
    # DISPLAY RESULTS
    # ======================================
    st.success(f"✅ Matching complete! Found {len(matched)} matched invoices.")
    st.subheader("🟩 Matched Invoices")
    st.dataframe(matched)

    st.subheader("🟥 Missing in ERP")
    st.dataframe(erp_missing)

    st.subheader("🟦 Missing in Vendor")
    st.dataframe(ven_missing)
