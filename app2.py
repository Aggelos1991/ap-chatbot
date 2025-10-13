import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

# ======================================
# CONFIG
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ======================================
# HELPER FUNCTIONS
# ======================================

def normalize_number(v):
    """Safely convert values like '1.234,56', '1,234.56', or Series to float."""
    import pandas as pd

    if isinstance(v, (pd.Series, list)):
        v = v.iloc[0] if isinstance(v, pd.Series) else v[0]
    if v is None or (isinstance(v, float) and pd.isna(v)):
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
    """Map multilingual headers to unified names for ERP or Vendor."""
    mapping = {
        "invoice": [
            "alternative document", "alt document", "invoice", "factura",
            "nº factura", "document", "παραστατικό"
        ],
        "credit": [
            "credit", "crédito", "haber", "importe", "total", "valor"
        ],
        "debit": [
            "debit", "debe", "cargo", "χρέωση"
        ],
        "date": [
            "date", "fecha", "ημερομηνία"
        ],
        "vendor": [
            "supplier", "proveedor", "vendor", "προμηθευτής"
        ],
        "trn": [
            "vat", "cif", "trn", "afm", "tax id"
        ],
        "description": [
            "description", "descripción", "concepto", "περιγραφή"
        ]
    }

    rename_map = {}
    for k, vals in mapping.items():
        for col in df.columns:
            if any(v in str(col).lower() for v in vals):
                rename_map[col] = f"{k}_{tag}"

    df = df.rename(columns=rename_map)
    return df


def extract_core_invoice(inv):
    """Extracts main numeric/alphanumeric part from invoice numbers."""
    if not inv or pd.isna(inv):
        return ""
    s = str(inv).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    match = re.search(r"([A-Z]*\d{2,6})$", s)
    return match.group(1) if match else s[-4:] if len(s) > 4 else s


# ======================================
# CORE MATCHING LOGIC
# ======================================

def match_invoices(erp_df, ven_df):
    matched, matched_erp, matched_ven = [], set(), set()

    # --- Clean and focus on essential columns ---
    if "invoice_erp" not in erp_df.columns:
        st.error("❌ 'Alternative Document' (invoice) not found in ERP file.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    if "credit_erp" not in erp_df.columns and "debit_erp" not in erp_df.columns:
        st.error("❌ 'Credit' or 'Charge' column not found in ERP file.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Remove payments, transfers, etc.
    payment_words = ["pago", "payment", "transfer", "bank", "liquidación"]
    for df, tag in [(erp_df, "erp"), (ven_df, "ven")]:
        desc_col = f"description_{tag}"
        if desc_col in df.columns:
            df.drop(
                df[df[desc_col].astype(str).str.lower().apply(
                    lambda x: any(k in x for k in payment_words)
                )].index,
                inplace=True
            )

    # Extract core invoice IDs
    erp_df["__core"] = erp_df["invoice_erp"].astype(str).apply(extract_core_invoice)
    if "invoice_ven" in ven_df.columns:
        ven_df["__core"] = ven_df["invoice_ven"].astype(str).apply(extract_core_invoice)
    else:
        st.warning("⚠️ No invoice column found in Vendor file — matching will be limited.")
        ven_df["__core"] = ""

    # --- Matching loop ---
    for _, e_row in erp_df.iterrows():
        e_inv = str(e_row.get("invoice_erp", "")).strip()
        if not e_inv:
            continue

        credit_val = e_row.get("credit_erp", 0)
        debit_val = e_row.get("debit_erp", 0)
        e_amt = normalize_number(credit_val) or -normalize_number(debit_val)
        e_core = e_row["__core"]

        for _, v_row in ven_df.iterrows():
            v_inv = str(v_row.get("invoice_ven", "")).strip()
            if not v_inv:
                continue

            v_core = v_row["__core"]
            desc = str(v_row.get("description_ven", "")).lower()
            v_amt = normalize_number(v_row.get("credit_ven", 0)) or normalize_number(v_row.get("amount_ven", 0))

            # Smart match: exact, fuzzy, or partial numeric
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
                    "Invoice (ERP)": e_inv,
                    "Invoice (Vendor)": v_inv,
                    "ERP Amount (Credit)": e_amt,
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


# ======================================
# STREAMLIT UI
# ======================================

st.write("Upload your ERP Export (Alternative Document / Credit / Charge) and Vendor Statement (Factura / Crédito / Importe):")

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

    # --- Highlighting ---
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        else:
            return [''] * len(row)

    # --- Display tables ---
    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1))
    else:
        st.info("No matches or differences found.")

    st.subheader("❌ Missing in ERP")
    if not erp_missing.empty and len(erp_missing.columns) > 0:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"))
    else:
        st.success("✅ No missing invoices in ERP file.")

    st.subheader("❌ Missing in Vendor")
    if not ven_missing.empty and len(ven_missing.columns) > 0:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"))
    else:
        st.success("✅ No missing invoices in Vendor file.")

    # --- Downloads ---
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
