import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ==========================
# Helper functions
# ==========================
def normalize_number(v):
    """Safely convert strings like '1.234,56' or '1,234.56' to float, handling Series and NaN."""
    import pandas as pd

    # Handle Series or list values
    if isinstance(v, (pd.Series, list)):
        v = v.iloc[0] if isinstance(v, pd.Series) else v[0]

    # Handle NaN or None
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0

    # Convert to string, clean symbols and parse
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
    except Exception:
        return 0.0


def normalize_columns(df, tag):
    """Unify multilingual headers (ERP/Vendor)."""
    mapping = {
        "vendor": ["supplier", "proveedor", "vendor", "προμηθευτής"],
        "trn": ["vat", "cif", "trn", "afm", "tax id"],
        "invoice": ["invoice", "alt document", "alternative document", "factura", "παραστατικό"],
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


def extract_core_invoice(inv):
    """Extract the meaningful part of the invoice (last alphanumeric tail)."""
    if not inv or pd.isna(inv):
        return ""
    s = str(inv).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    match = re.search(r"([A-Z]*\d{2,5})$", s)
    return match.group(1) if match else s[-4:] if len(s) > 4 else s


# ==========================
# Core reconciliation
# ==========================
def match_invoices(erp_df, ven_df):
    """Match invoices & credit notes — ignore all payments."""
    matched, matched_erp, matched_ven = [], set(), set()

    # Detect TRN (VAT/CIF) and limit ERP scope to this vendor only
    trn_col_erp = next((c for c in erp_df.columns if "trn_" in c), None)
    trn_col_ven = next((c for c in ven_df.columns if "trn_" in c), None)
    if trn_col_erp and trn_col_ven:
        vendor_trn = str(ven_df[trn_col_ven].dropna().iloc[0]).strip()
        erp_df = erp_df[erp_df[trn_col_erp].astype(str).str.strip() == vendor_trn]

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

    # Extract core invoice part for smart matching
    erp_df["__core"] = erp_df["invoice_erp"].astype(str).apply(extract_core_invoice)
    ven_df["__core"] = ven_df["invoice_ven"].astype(str).apply(extract_core_invoice)

    # Match logic
    for _, e_row in erp_df.iterrows():
        e_inv = str(e_row.get("invoice_erp", "")).strip()
        if not e_inv:
            continue

        # SAFER extraction for credit/debit amounts
        credit_val = e_row["credit_erp"] if "credit_erp" in e_row else 0
        debit_val = e_row["debit_erp"] if "debit_erp" in e_row else 0
        if isinstance(credit_val, pd.Series):
            credit_val = credit_val.iloc[0]
        if isinstance(debit_val, pd.Series):
            debit_val = debit_val.iloc[0]
        e_amt = normalize_number(credit_val) or -normalize_number(debit_val)

        e_core = e_row["__core"]

        for _, v_row in ven_df.iterrows():
            v_inv = str(v_row.get("invoice_ven", "")).strip()
            if not v_inv:
                continue
            v_core = v_row["__core"]
            desc = str(v_row.get("description_ven", "")).lower()
            d_val = normalize_number(v_row.get("debit_ven", 0))
            c_val = normalize_number(v_row.get("credit_ven", 0))
            v_amt = d_val if "abono" not in desc and "credit" not in desc else -c_val

            # Smart match: exact, fuzzy, or suffix match
            if (
                e_inv == v_inv
                or e_core == v_core
                or v_core.endswith(e_core)
                or e_core.endswith(v_core)
                or fuzz.ratio(e_inv, v_inv) > 90
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
        "⬇️ Download Matched/Differences CSV",
        matched.to_csv(index=False).encode("utf-8"),
        "reconciliation_results.csv",
        "text/csv"
    )
    st.download_button(
        "⬇️ Download Missing in ERP CSV",
        erp_missing.to_csv(index=False).encode("utf-8"),
        "missing_in_erp.csv",
        "text/csv"
    )
    st.download_button(
        "⬇️ Download Missing in Vendor CSV",
        ven_missing.to_csv(index=False).encode("utf-8"),
        "missing_in_vendor.csv",
        "text/csv"
    )
else:
    st.info("Please upload both ERP and Vendor files to begin.")
