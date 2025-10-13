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
# HELPERS
# ======================================
def normalize_number(v):
    """Convert numeric strings like '1.234,56' or '1,234.56' safely to float."""
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
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "alternative document", "alt document", "invoice", "factura",
            "nº factura", "document", "παραστατικό"
        ],
        "credit": ["credit", "crédito", "haber"],
        "debit": ["debit", "debe", "cargo", "total", "importe", "valor"],
    }
    rename_map = {}
    for k, vals in mapping.items():
        for col in df.columns:
            if any(v in str(col).lower() for v in vals):
                rename_map[col] = f"{k}_{tag}"
    return df.rename(columns=rename_map)


def extract_core_invoice(inv):
    """Extract meaningful invoice tail."""
    import pandas as pd
    if isinstance(inv, (pd.Series, list)):
        inv = inv.iloc[0] if isinstance(inv, pd.Series) else inv[0]
    if inv is None or (isinstance(inv, float) and pd.isna(inv)):
        return ""
    s = str(inv).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    match = re.search(r"([A-Z]*\d{2,6})$", s)
    return match.group(1) if match else (s[-4:] if len(s) > 4 else s)

# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched, matched_erp, matched_ven = [], set(), set()

    # Validate essential columns
    if "invoice_erp" not in erp_df.columns:
        st.error("❌ 'Alternative Document' not found in ERP file.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if not ("credit_erp" in erp_df.columns or "debit_erp" in erp_df.columns):
        st.error("❌ 'Credit' or 'Charge' column not found in ERP file.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if "invoice_ven" not in ven_df.columns:
        st.error("❌ 'Invoice / Factura' not found in Vendor file.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Extract simplified invoice codes
    erp_df["__core"] = erp_df["invoice_erp"].astype(str).apply(extract_core_invoice)
    ven_df["__core"] = ven_df["invoice_ven"].astype(str).apply(extract_core_invoice)

    for _, e_row in erp_df.iterrows():
        e_inv = str(e_row["invoice_erp"]).strip()
        e_core = e_row["__core"]
        e_amt = normalize_number(e_row.get("credit_erp", 0)) or -normalize_number(e_row.get("debit_erp", 0))

        for _, v_row in ven_df.iterrows():
            v_inv = str(v_row["invoice_ven"]).strip()
            v_core = v_row["__core"]

            # Vendor amount (priority: debit/debe/importe/total, fallback to credit)
            v_amt = (
                normalize_number(v_row.get("debit_ven", 0))
                or normalize_number(v_row.get("credit_ven", 0))
            )

            # Matching logic: exact, suffix, or fuzzy
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
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": status
                })
                matched_erp.add(e_inv)
                matched_ven.add(v_inv)
                break

    # Filter missing invoices (exclude already matched)
    erp_missing = (
        erp_df[~erp_df["invoice_erp"].isin(matched_erp)]
        .loc[:, ["invoice_erp", "credit_erp", "debit_erp"]]
        .rename(columns={
            "invoice_erp": "ERP Invoice",
            "credit_erp": "ERP Amount (Credit)",
            "debit_erp": "ERP Amount (Debit)"
        })
        .reset_index(drop=True)
    )

    ven_missing = (
        ven_df[~ven_df["invoice_ven"].isin(matched_ven)]
        .loc[:, ["invoice_ven", "debit_ven", "credit_ven"]]
        .rename(columns={
            "invoice_ven": "Vendor Invoice",
            "debit_ven": "Vendor Amount (Debit)",
            "credit_ven": "Vendor Amount (Credit)"
        })
        .reset_index(drop=True)
    )

    df_matched = pd.DataFrame(matched)
    return df_matched, erp_missing, ven_missing

# ======================================
# STREAMLIT UI
# ======================================
st.write("Upload your ERP Export (Alternative Document / Credit / Charge) and Vendor Statement (Factura / Débito / Total):")

uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(pd.read_excel(uploaded_erp), "erp")
    ven_df = normalize_columns(pd.read_excel(uploaded_vendor), "ven")

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    # --- Summary ---
    total_match = len(matched[matched["Status"] == "Match"]) if not matched.empty else 0
    total_diff = len(matched[matched["Status"] == "Difference"]) if not matched.empty else 0
    total_missing = len(erp_missing) + len(ven_missing)
    st.success(f"✅ Recon complete: {total_match} matched, {total_diff} differences, {total_missing} missing")

    # --- Style functions ---
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        else:
            return [''] * len(row)

    # --- Display clean tables ---
    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1))
    else:
        st.info("No matches or differences found.")

    st.subheader("❌ Missing in ERP")
    if not erp_missing.empty:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"))
    else:
        st.success("✅ No missing invoices in ERP file.")

    st.subheader("❌ Missing in Vendor")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"))
    else:
        st.success("✅ No missing invoices in Vendor file.")

    # --- Downloads ---
    st.download_button("⬇️ Matched/Differences CSV", matched.to_csv(index=False).encode("utf-8"), "matched_results.csv", "text/csv")
    st.download_button("⬇️ Missing in ERP CSV", erp_missing.to_csv(index=False).encode("utf-8"), "missing_in_erp.csv", "text/csv")
    st.download_button("⬇️ Missing in Vendor CSV", ven_missing.to_csv(index=False).encode("utf-8"), "missing_in_vendor.csv", "text/csv")

else:
    st.info("Please upload both ERP and Vendor files to begin.")
