import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

# ================== APP SETUP ==================
st.set_page_config(page_title="🦖 ReconRaptor", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

st.markdown("""
Upload your **ERP Export** and **Vendor Statement** Excel files.  
ReconRaptor automatically matches invoices & credit notes by TRN and amount.
""")

# ================== FILE UPLOAD ==================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

# ================== SAFE READER ==================
def safe_read(file, label):
    try:
        df = pd.read_excel(file)
        st.success(f"✅ {label} file loaded ({len(df)} rows)")
        st.dataframe(df.head(), use_container_width=True)
        return df
    except Exception as e:
        st.error(f"❌ Could not read {label} file: {e}")
        return pd.DataFrame()

# ================== NORMALIZE COLUMNS ==================
def normalize_columns(df, source="erp"):
    rename_map = {}

    mapping_patterns = {
        "vendor": ["supplier", "proveedor", "vendor", "cliente"],
        "trn": ["vat", "cif", "afm", "trn", "tax id"],
        "invoice": ["invoice number", "factura", "inv", "document", "παραστατικό"],
        "amount": ["amount", "importe", "valor"],
        "balance": ["balance", "saldo"],
        "description": ["description", "descripción", "περιγραφή"],
        "debit": ["debit", "debe", "χρέωση"],
        "credit": ["credit", "haber", "πίστωση"],
        "date": ["date", "fecha", "ημερομηνία"]
    }

    for key, patterns in mapping_patterns.items():
        for col in df.columns:
            c = str(col).strip().lower()
            if any(p in c for p in patterns):
                rename_map[col] = f"{key}_{source}"
                break

    return df.rename(columns=rename_map)

# ================== NUMBER CLEANER ==================
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

# ================== MATCH LOGIC ==================
def match_invoices(erp_df, ven_df):
    try:
        erp_df = erp_df.reset_index().rename(columns={"index": "_id_erp"})
        ven_df = ven_df.reset_index().rename(columns={"index": "_id_ven"})

        matched, matched_erp, matched_ven = [], set(), set()

        # detect vendor TRN
        trn_col = next((c for c in ven_df.columns if "trn_" in c), None)
        if trn_col:
            vendor_trn = str(ven_df[trn_col].dropna().iloc[0])
            erp_df = erp_df[erp_df.get("trn_erp", "") == vendor_trn]

        for _, e in erp_df.iterrows():
            e_inv = str(e.get("invoice_erp", "")).strip()
            e_amt = normalize_number(e.get("amount_erp", 0))
            e_desc = str(e.get("description_erp", "")).lower()

            # skip zero or blank invoice numbers
            if not e_inv:
                continue

            for _, v in ven_df.iterrows():
                v_inv = str(v.get("invoice_ven", "")).strip()
                v_desc = str(v.get("description_ven", "")).lower()

                # Skip payments (Pago parcial, Payment)
                if any(x in v_desc for x in ["pago parcial", "payment", "transfer"]):
                    continue

                # Determine amount (Debit = invoice, Credit = CN)
                d_val = normalize_number(v.get("debit_ven", 0))
                c_val = normalize_number(v.get("credit_ven", 0))
                v_amt = c_val if "abono" in v_desc or "credit" in v_desc else d_val

                if e_inv[-4:] in v_inv or fuzz.ratio(e_inv, v_inv) > 80:
                    diff = round(e_amt - v_amt, 2)
                    status = "Match" if abs(diff) < 0.05 else "Difference"

                    matched.append({
                        "Vendor/Supplier": e.get("vendor_erp", ""),
                        "TRN/AFM": e.get("trn_erp", ""),
                        "ERP Invoice": e_inv,
                        "Vendor Invoice": v_inv,
                        "ERP Amount": e_amt,
                        "Vendor Amount": v_amt,
                        "Difference": diff,
                        "Status": status,
                        "Description": v_desc
                    })
                    matched_erp.add(e_inv)
                    matched_ven.add(v_inv)
                    break

        df_matched = pd.DataFrame(matched)
        matched_invoices = df_matched["ERP Invoice"].unique().tolist()
        matched_vendor = df_matched["Vendor Invoice"].unique().tolist()

        erp_missing = erp_df[~erp_df["invoice_erp"].astype(str).isin(matched_invoices)].reset_index(drop=True)
        ven_missing = ven_df[~ven_df["invoice_ven"].astype(str).isin(matched_vendor)].reset_index(drop=True)

        return df_matched, erp_missing, ven_missing

    except Exception as e:
        st.error(f"❌ Matching error: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ================== COLOR HIGHLIGHTS ==================
def color_rows(row):
    if row["Status"] == "Match":
        return ['background-color: #2e7d32; color: white'] * len(row)  # green
    elif row["Status"] == "Difference":
        return ['background-color: #f9a825; color: black'] * len(row)  # yellow
    else:
        return [''] * len(row)

# ================== MAIN EXECUTION ==================
if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(safe_read(uploaded_erp, "ERP Export"), "erp")
    ven_df = normalize_columns(safe_read(uploaded_vendor, "Vendor Statement"), "ven")

    if not erp_df.empty and not ven_df.empty:
        with st.spinner("Reconciling invoices... please wait"):
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

        st.success(f"✅ Recon complete: {len(matched)} matched, {len(erp_missing)} missing in ERP, {len(ven_missing)} missing in Vendor")

        # Show matched/differences
        if not matched.empty:
            st.subheader("📊 Matched / Differences")
            st.dataframe(matched.style.apply(color_rows, axis=1), use_container_width=True)

        # Show missing in ERP
        if not erp_missing.empty:
            st.subheader("❌ Missing in ERP")
            st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)

        # Show missing in Vendor
        if not ven_missing.empty:
            st.subheader("❌ Missing in Vendor")
            st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)

        st.download_button(
            "⬇️ Download Matched CSV",
            matched.to_csv(index=False).encode("utf-8"),
            "ReconRaptor_Results.csv",
            "text/csv"
        )

else:
    st.info("Please upload both Excel files to begin.")
