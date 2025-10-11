import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

# === APP SETUP ===
st.set_page_config(page_title="🦖 ReconRaptor", layout="wide")
st.title("🦖 ReconRaptor")

st.markdown("""
Upload both your **ERP Export** and **Vendor Statement** Excel files below.  
ReconRaptor will automatically match invoices and credit notes by TRN and amount.
""")

# === FILE UPLOADS ===
uploaded_erp = st.file_uploader("📂 Upload ERP Export", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement", type=["xlsx"])

# === SAFE READER ===
def safe_read(file, label):
    try:
        df = pd.read_excel(file)
        st.success(f"✅ {label} file loaded ({len(df)} rows)")
        st.write(df.head())
        return df
    except Exception as e:
        st.error(f"❌ Failed to read {label} file: {e}")
        return pd.DataFrame()

# === HELPER FUNCTIONS ===
def normalize_number(value):
    if pd.isna(value):
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(value))
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

# === MATCHING FUNCTION ===
def match_invoices(erp_df, ven_df):
    try:
        # Reset
        erp_df = erp_df.reset_index().rename(columns={"index": "_id_erp"})
        ven_df = ven_df.reset_index().rename(columns={"index": "_id_ven"})

        # Guess TRN if present
        selected_trn = None
        if "trn_ven" in ven_df.columns and not ven_df["trn_ven"].dropna().empty:
            selected_trn = str(ven_df["trn_ven"].dropna().iloc[0]).strip()

        # Filter ERP by same TRN
        if selected_trn and "trn_erp" in erp_df.columns:
            erp_df = erp_df[erp_df["trn_erp"].astype(str).str.strip() == selected_trn]

        # Simple match logic
        matches = []
        for _, e in erp_df.iterrows():
            e_inv = str(e.get("invoice_erp", "")).strip()
            e_amt = normalize_number(e.get("amount_erp", 0))
            for _, v in ven_df.iterrows():
                v_inv = str(v.get("invoice_ven", "")).strip()
                v_amt = normalize_number(v.get("debit_ven", 0) or v.get("credit_ven", 0))
                if e_inv[-5:] in v_inv or fuzz.ratio(e_inv, v_inv) > 85:
                    diff = round(e_amt - v_amt, 2)
                    matches.append({
                        "ERP Invoice": e_inv,
                        "Vendor Invoice": v_inv,
                        "ERP Amount": e_amt,
                        "Vendor Amount": v_amt,
                        "Difference": diff,
                        "Status": "Match" if abs(diff) < 0.05 else "Difference"
                    })
                    break

        matched_df = pd.DataFrame(matches)
        matched_invoices = matched_df["ERP Invoice"].tolist()
        erp_missing = erp_df[~erp_df["invoice_erp"].astype(str).isin(matched_invoices)]
        ven_missing = ven_df[~ven_df["invoice_ven"].astype(str).isin(matched_df["Vendor Invoice"].tolist())]

        return matched_df, erp_missing, ven_missing
    except Exception as e:
        st.error(f"❌ Matching error: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# === RUN LOGIC ===
if uploaded_erp and uploaded_vendor:
    erp_df = safe_read(uploaded_erp, "ERP Export")
    ven_df = safe_read(uploaded_vendor, "Vendor Statement")

    if not erp_df.empty and not ven_df.empty:
        with st.spinner("Reconciling invoices..."):
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

        if matched.empty and erp_missing.empty and ven_missing.empty:
            st.warning("⚠️ No matches found. Please check your column names.")
        else:
            st.success(f"✅ Recon complete: {len(matched)} matched, {len(erp_missing)} missing in ERP, {len(ven_missing)} missing in Vendor")

            st.subheader("📊 Matched Invoices")
            st.dataframe(matched, use_container_width=True)

            st.subheader("❌ Missing in ERP")
            st.dataframe(erp_missing, use_container_width=True)

            st.subheader("❌ Missing in Vendor")
            st.dataframe(ven_missing, use_container_width=True)

else:
    st.info("Please upload both Excel files to begin.")
