import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

st.set_page_config(page_title="⚙️ ReconRaptor", layout="wide")
st.title("🦖 ReconRaptor — Vendor Reconciliation (Invoices & Credit Notes Only)")

# ============================================================
# Helper: Normalize numeric strings (EU/US formats)
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
# Helper: Normalize multilingual column names
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
        "balance": ["balance", "saldo", "υπόλοιπο"],
        "date": ["date", "fecha", "ημερομηνία"]
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
# Core Matching Logic
# ============================================================
def match_invoices(erp_df, ven_df):
    erp_df = erp_df.reset_index().rename(columns={"index": "_id_erp"})
    ven_df = ven_df.reset_index().rename(columns={"index": "_id_ven"})

    # 🧹 Clean vendor data: drop payments, transfers, etc.
    skip_words = [
        "pago", "transferencia", "payment", "paid", "bank",
        "deposit", "wire", "transf", "πληρωμή"
    ]
    ven_df = ven_df[
        ~ven_df["description_ven"].astype(str).str.lower().apply(
            lambda x: any(w in x for w in skip_words)
        )
    ].reset_index(drop=True)

    matched = []
    matched_erp, matched_ven = set(), set()

    vendor_trn = str(ven_df["trn_ven"].iloc[0]) if "trn_ven" in ven_df.columns else None
    if vendor_trn:
        erp_df = erp_df[erp_df["trn_erp"] == vendor_trn]

    for _, e_row in erp_df.iterrows():
        e_trn = str(e_row.get("trn_erp", "")).strip()
        e_inv = str(e_row.get("invoice_erp", "")).strip()
        e_amt = normalize_number(e_row.get("amount_erp", e_row.get("credit_erp", 0)))
        e_id = e_row["_id_erp"]

        ven_subset = ven_df[ven_df["trn_ven"] == e_trn]
        for _, v_row in ven_subset.iterrows():
            v_id = v_row["_id_ven"]
            if v_id in matched_ven:
                continue

            v_inv = str(v_row.get("invoice_ven", "")).strip()
            desc = str(v_row.get("description_ven", "")).lower()
            d_val = normalize_number(v_row.get("debit_ven", 0))
            c_val = normalize_number(v_row.get("credit_ven", 0))

            # Vendor debit = invoice, credit = credit note
            if any(w in desc for w in ["abono", "credit"]):
                v_amt = c_val
            else:
                v_amt = d_val

            # Flexible invoice match (by last digits or fuzzy ratio)
            if (
                e_inv[-4:] in v_inv
                or v_inv[-4:] in e_inv
                or fuzz.ratio(e_inv, v_inv) > 78
            ):
                diff = round(e_amt - v_amt, 2)
                status = "Match" if diff == 0 else "Difference"

                matched.append({
                    "Vendor/Supplier": e_row.get("vendor_erp", ""),
                    "TRN/AFM": e_trn,
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": status,
                    "Description": desc
                })
                matched_erp.add(e_id)
                matched_ven.add(v_id)
                break

    # ✅ Avoid double-flagging
    df_matched = pd.DataFrame(matched)
    matched_invoices = df_matched["ERP Invoice"].unique().tolist()
    matched_vendor_invoices = df_matched["Vendor Invoice"].unique().tolist()

    erp_missing = erp_df[~erp_df["invoice_erp"].astype(str).isin(matched_invoices)].reset_index(drop=True)
    ven_missing = ven_df[~ven_df["invoice_ven"].astype(str).isin(matched_vendor_invoices)].reset_index(drop=True)

    return df_matched, erp_missing, ven_missing

# ============================================================
# Streamlit Interface
# ============================================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(pd.read_excel(uploaded_erp), "erp")
    ven_df = normalize_columns(pd.read_excel(uploaded_vendor), "ven")

    with st.spinner("ReconRaptor scanning invoices... 🦖"):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    total_m = len(matched)
    total_d = len(matched[matched["Status"] == "Difference"])
    total_miss = len(erp_missing) + len(ven_missing)
    st.success(f"✅ Recon complete: {total_m} Matches · {total_d} Differences · {total_miss} Missing")

    # --- Highlighting
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)  # green
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)  # yellow
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
        "ReconRaptor_Results.csv",
        "text/csv"
    )

else:
    st.info("Please upload both ERP Export and Vendor Statement files to begin the hunt. 🦖")
