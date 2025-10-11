import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

st.set_page_config(page_title="🦅 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦅 ReconRaptor — Vendor Invoice Reconciliation")

# ==========================
# Helper functions
# ==========================
def normalize_number(v):
    """Convert strings like '1.234,56' or '1,234.56' to float."""
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
    """Unify multilingual headers (ERP/Vendor)."""
    mapping = {
        "vendor": ["supplier", "proveedor", "vendor", "προμηθευτής"],
        "trn": ["vat", "cif", "trn", "afm", "tax id"],
        "invoice": ["invoice", "alt document", "factura", "παραστατικό"],
        "description": ["description", "descripción", "περιγραφή"],
        "debit": ["debit", "debe", "χρέωση"],
        "credit": ["credit", "haber", "πίστωση"],
        "amount": ["amount", "importe", "valor"],
        "balance": ["balance", "saldo", "υπόλοιπο"],
        "date": ["date", "fecha", "ημερομηνία"],
    }
    rename_map = {}
    for k, vals in mapping.items():
        for col in df.columns:
            if any(v in str(col).lower() for v in vals):
                rename_map[col] = f"{k}_{tag}"
    return df.rename(columns=rename_map)


# ==========================
# Core reconciliation
# ==========================
def match_invoices(erp_df, ven_df):
    """Match invoices & credit notes — ignore all payments."""
    matched, matched_erp, matched_ven = [], set(), set()

    # Detect TRN and limit scope to same vendor
    trn_col_erp = next((c for c in erp_df.columns if "trn_" in c), None)
    trn_col_ven = next((c for c in ven_df.columns if "trn_" in c), None)
    if trn_col_erp and trn_col_ven:
        vendor_trn = str(ven_df[trn_col_ven].dropna().iloc[0])
        erp_df = erp_df[erp_df[trn_col_erp].astype(str) == vendor_trn]

    # Remove payment/transfer lines
    payment_words = ["pago", "payment", "transfer", "bank", "liquidación", "partial"]
    for df, desc_col in [(erp_df, "description_erp"), (ven_df, "description_ven")]:
        if desc_col in df.columns:
            df.drop(
                df[df[desc_col].astype(str).str.lower().apply(
                    lambda x: any(k in x for k in payment_words)
                )].index,
                inplace=True,
            )

    # Match logic
    for _, e_row in erp_df.iterrows():
        e_inv = str(e_row.get("invoice_erp", "")).strip()
        if not e_inv:
            continue
        e_amt = normalize_number(e_row.get("amount_erp", 0))

        for _, v_row in ven_df.iterrows():
            v_inv = str(v_row.get("invoice_ven", "")).strip()
            if not v_inv:
                continue
            desc = str(v_row.get("description_ven", "")).lower()
            d_val = normalize_number(v_row.get("debit_ven", 0))
            c_val = normalize_number(v_row.get("credit_ven", 0))
            v_amt = c_val if "abono" in desc or "credit" in desc else d_val

            # ✅ Strict matching: only match identical or 92%+ similar invoice numbers
            if (
                e_inv == v_inv
                or fuzz.ratio(e_inv, v_inv) > 92
                or e_inv.replace(".", "").replace("-", "") == v_inv.replace(".", "").replace("-", "")
            ):
                diff = round(e_amt - v_amt, 2)
                status = "Match" if abs(diff) < 0.05 else "Difference"
                matched.append({
                    "Vendor/Supplier": e_row.get("vendor_erp", ""),
                    "TRN/AFM": e_row.get("trn_erp", ""),
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": status,
                    "Description": desc
                })
                matched_erp.add(e_inv)
                matched_ven.add(v_inv)
                break

    df_matched = pd.DataFrame(matched)
    erp_missing = erp_df[~erp_df["invoice_erp"].isin(matched_erp)].reset_index(drop=True)
    ven_missing = ven_df[~ven_df["invoice_ven"].isin(matched_ven)].reset_index(drop=True)
    return df_matched, erp_missing, ven_missing


# ==========================
# Streamlit UI
# ==========================
st.write("Upload your ERP Export and Vendor Statement for reconciliation:")

uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(pd.read_excel(uploaded_erp), "erp")
    ven_df = normalize_columns(pd.read_excel(uploaded_vendor), "ven")

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    total_match = len(matched[matched["Status"] == "Match"]) if not matched.empty else 0
    total_diff = len(matched[matched["Status"] == "Difference"]) if not matched.empty else 0
    total_missing = len(erp_missing) + len(ven_missing)
    st.success(f"✅ Recon complete: {total_match} matched, {total_diff} differences, {total_missing} missing")

    # Highlight rows
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
        "reconciliation_results.csv",
        "text/csv"
    )
else:
    st.info("Please upload both ERP and Vendor files to begin.")
