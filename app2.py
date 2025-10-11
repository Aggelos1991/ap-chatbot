import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

st.set_page_config(page_title="🤝 Vendor Reconciliation", layout="wide")
st.title("🧾 Vendor Reconciliation — Accurate Differences")

# ============================================================
# Helper: Normalize numbers (EU, US formats)
# ============================================================
def normalize_number(value):
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
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


# ============================================================
# Column mapping — multilingual
# ============================================================
def normalize_columns(df, source="ven"):
    colmap = {
        "vendor": ["supplier name", "vendor", "proveedor", "προμηθευτής"],
        "trn": ["tax id", "cif", "vat", "afm", "trn", "vat number"],
        "invoice": ["invoice number", "alt document", "invoice", "factura", "παραστατικό"],
        "description": ["description", "descripción", "περιγραφή"],
        "debit": ["debit", "debe", "χρέωση"],
        "credit": ["credit", "haber", "πίστωση"],
        "amount": ["amount", "importe", "valor"],
    }

    rename_map = {}
    for key, variants in colmap.items():
        for col in df.columns:
            c = str(col).strip().lower()
            if any(v in c for v in variants):
                rename_map[col] = f"{key}_{source}"
                break
    return df.rename(columns=rename_map)


# ============================================================
# Matching Logic — clean version
# ============================================================
def match_invoices(erp_df, ven_df):
    matched, matched_erp, matched_ven = [], set(), set()

    vendor_trn = str(ven_df["trn_ven"].iloc[0]) if "trn_ven" in ven_df.columns else None
    if vendor_trn:
        erp_df = erp_df[erp_df["trn_erp"] == vendor_trn]

    for e_idx, e_row in erp_df.iterrows():
        e_trn = str(e_row.get("trn_erp", "")).strip()
        e_inv = str(e_row.get("invoice_erp", "")).strip()
        e_amt = normalize_number(e_row.get("amount_erp", e_row.get("credit_erp", 0)))

        ven_subset = ven_df[ven_df["trn_ven"] == e_trn]
        for v_idx, v_row in ven_subset.iterrows():
            v_inv = str(v_row.get("invoice_ven", "")).strip()
            desc = str(v_row.get("description_ven", "")).lower()
            d_val = normalize_number(v_row.get("debit_ven", 0))
            c_val = normalize_number(v_row.get("credit_ven", 0))

            if "pago" in desc or "transferencia" in desc or "payment" in desc:
                continue  # ignore payments

            v_amt = c_val if ("abono" in desc or "credit" in desc) else d_val

            # flexible matching
            if (
                e_inv[-4:] in v_inv
                or v_inv[-4:] in e_inv
                or fuzz.ratio(e_inv, v_inv) > 75
            ):
                diff = round(e_amt - v_amt, 2)
                matched.append({
                    "Vendor/Supplier": e_row.get("vendor_erp", ""),
                    "TRN/AFM": e_trn,
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match" if diff == 0 else "Difference",
                    "Description": desc
                })
                matched_erp.add(e_idx)
                matched_ven.add(v_idx)
                break

    erp_missing = erp_df.loc[~erp_df.index.isin(matched_erp)].reset_index(drop=True)
    ven_missing = ven_df.loc[~ven_df.index.isin(matched_ven)].reset_index(drop=True)
    return pd.DataFrame(matched), erp_missing, ven_missing


# ============================================================
# Streamlit Interface
# ============================================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(pd.read_excel(uploaded_erp), "erp")
    ven_df = normalize_columns(pd.read_excel(uploaded_vendor), "ven")

    with st.spinner("Reconciling... please wait"):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    st.success("✅ Reconciliation complete!")
    st.subheader("📊 Matched / Differences")
    st.dataframe(matched)

    st.subheader("❌ Missing in ERP")
    st.dataframe(erp_missing)

    st.subheader("❌ Missing in Vendor")
    st.dataframe(ven_missing)

    st.download_button("⬇️ Download Matched CSV",
        matched.to_csv(index=False).encode("utf-8"),
        "matched_results.csv",
        "text/csv")
else:
    st.info("Please upload both ERP Export and Vendor Statement files to begin.")
