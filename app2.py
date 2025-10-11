import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
import re

st.set_page_config(page_title="🦖 ReconRaptor", layout="wide")
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

    # ---------- 1️⃣ TRN / Vendor filtering ----------
    if "trn_ven" not in ven_df.columns or ven_df["trn_ven"].dropna().empty:
        return pd.DataFrame([]), erp_df.iloc[0:0], ven_df.iloc[0:0]

    selected_trn = str(ven_df["trn_ven"].dropna().iloc[0]).strip()
    ven_df = ven_df[ven_df["trn_ven"].astype(str).str.strip() == selected_trn].copy()
    erp_df = erp_df[erp_df["trn_erp"].astype(str).str.strip() == selected_trn].copy()

    # ---------- 2️⃣ Ignore payments ----------
    skip_words = ["pago", "transferencia", "payment", "paid", "bank", "deposit", "wire", "transf", "πληρωμή"]
    if "description_ven" in ven_df.columns:
        ven_df = ven_df[
            ~ven_df["description_ven"].astype(str).str.lower().apply(
                lambda x: any(w in x for w in skip_words)
            )
        ].reset_index(drop=True)

    # ---------- Helpers ----------
    def is_cn_vendor(row):
        desc = str(row.get("description_ven", "")).lower()
        credit = normalize_number(row.get("credit_ven", 0))
        return ("abono" in desc or "credit" in desc) or credit > 0

    def vendor_amount(row):
        return normalize_number(row.get("credit_ven", 0)) if is_cn_vendor(row) else normalize_number(row.get("debit_ven", 0))

    def is_cn_erp(row):
        return normalize_number(row.get("amount_erp", row.get("credit_erp", 0))) < 0

    def erp_amount(row):
        return abs(normalize_number(row.get("amount_erp", row.get("credit_erp", 0))))

    def last_digits(s, k=6):
        s = str(s)
        digits = re.findall(r"\d+", s)
        if not digits:
            return ""
        return "".join(digits)[-k:]

    def invoice_match(a, b):
        ta, tb = last_digits(a), last_digits(b)
        for n in (6, 5, 4, 3):
            if len(ta) >= n and len(tb) >= n and ta[-n:] == tb[-n:]:
                return True
        return fuzz.ratio(str(a), str(b)) >= 90

    # ---------- 3️⃣ Matching ----------
    matched_rows = []
    used_ven_ids, used_erp_ids = set(), set()

    for _, e in erp_df.iterrows():
        e_id, e_inv = e["_id_erp"], str(e.get("invoice_erp", "")).strip()
        e_amt, e_iscn = abs(erp_amount(e)), is_cn_erp(e)

        for _, v in ven_df.iterrows():
            v_id = v["_id_ven"]
            if v_id in used_ven_ids:
                continue

            v_inv, v_amt, v_iscn = str(v.get("invoice_ven", "")).strip(), abs(vendor_amount(v)), is_cn_vendor(v)
            if e_iscn != v_iscn:
                continue
            if not invoice_match(e_inv, v_inv):
                continue

            diff = round(e_amt - v_amt, 2)
            status = "Match" if abs(diff) <= 0.01 else "Difference"

            matched_rows.append({
                "Vendor/Supplier": e.get("vendor_erp", ""),
                "TRN/AFM": selected_trn,
                "ERP Invoice": e_inv,
                "Vendor Invoice": v_inv,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status,
                "Description": str(v.get("description_ven", "")),
            })

            used_ven_ids.add(v_id)
            used_erp_ids.add(e_id)
            break

    matched_df = pd.DataFrame(matched_rows)
    matched_erp_invoices = set(matched_df["ERP Invoice"].astype(str)) if not matched_df.empty else set()
    matched_ven_invoices = set(matched_df["Vendor Invoice"].astype(str)) if not matched_df.empty else set()

    erp_missing = erp_df[~erp_df["invoice_erp"].astype(str).isin(matched_erp_invoices)].reset_index(drop=True)
    ven_missing = ven_df[~ven_df["invoice_ven"].astype(str).isin(matched_ven_invoices)].reset_index(drop=True)

    return matched_df, erp_missing, ven_missing

# ============================================================
# Streamlit Interface
# ============================================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(pd.read_excel(uploaded_erp), "erp")
    ven_df = normalize_columns(pd.read_excel(uploaded_vendor), "ven")

    with st.spinner("🦖 ReconRaptor analyzing invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    total_m = len(matched)
    total_d = len(matched[matched["Status"] == "Difference"])
    total_miss = len(erp_missing) + len(ven_missing)
    st.success(f"✅ Recon complete: {total_m} Matches · {total_d} Differences · {total_miss} Missing")

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
        "ReconRaptor_Results.csv",
        "text/csv"
    )
else:
    st.info("Please upload both ERP Export and Vendor Statement files to begin the hunt 🦖.")
