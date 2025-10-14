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
def normalize_number(v):
    """Convert numeric strings like '1.234,56' or '1,234.56' safely to float."""
    if v is None:
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
    """Map multilingual headers to unified names — optimized for Spanish vendor statements."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge",
            "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "descriptivo", "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento"
        ],
    }

    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    for k, vals in mapping.items():
        for col, low in cols_lower.items():
            if any(v in low for v in vals):
                rename_map[col] = f"{k}_{tag}"

    out = df.rename(columns=rename_map)

    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0

    return out


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    # ====== ERP PREP ======
    erp_df["__doctype"] = erp_df.apply(
        lambda r: "CN" if normalize_number(r.get("debit_erp")) > 0
        else ("INV" if normalize_number(r.get("credit_erp")) > 0 else "UNKNOWN"),
        axis=1
    )
    erp_df["__amt"] = erp_df.apply(
        lambda r: normalize_number(r["credit_erp"]) if r["__doctype"] == "INV"
        else (-normalize_number(r["debit_erp"]) if r["__doctype"] == "CN" else 0.0),
        axis=1
    )
    # ====== VENDOR PREP ======
    def detect_vendor_doc_type(row):
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))

        # ✅ Case 1: Credit Note explicitly in Credit column
        if credit > 0:
             return "CN"
        # ✅ Case 2: Negative debit also means a Credit Note
        elif debit < 0:
             return "CN"
        # ✅ Case 3: Normal invoice (positive debit)
        elif debit > 0:
             return "INV"
        else:
            return "UNKNOWN"


def calc_vendor_amount(row):
    debit = normalize_number(row.get("debit_ven"))
    credit = normalize_number(row.get("credit_ven"))
    doc = row.get("__doctype", "")

    if doc == "INV":
        # Normal invoice = positive amount (we owe the vendor)
        return abs(debit)
    elif doc == "CN":
        # Credit note = always negative impact
        return -abs(credit if credit > 0 else debit)
    else:
        return 0.0


# Apply to vendor dataframe
ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)
    

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # ====== MERGE ERP CREDIT/INVOICE PAIRS ======
    merged_rows = []
    grouped = erp_use.groupby("invoice_erp", dropna=False)

    for inv, group in grouped:
        if len(group) == 1:
            merged_rows.append(group.iloc[0])
            continue

        inv_rows = group[group["__doctype"] == "INV"]
        cn_rows = group[group["__doctype"] == "CN"]

        if not inv_rows.empty and not cn_rows.empty:
            total_inv = inv_rows["__amt"].sum()
            total_cn = cn_rows["__amt"].sum()
            net = round(total_inv + total_cn, 2)
            base_row = inv_rows.iloc[0].copy()
            base_row["__amt"] = net
            merged_rows.append(base_row)
        else:
            for _, row in group.iterrows():
                merged_rows.append(row)

    erp_use = pd.DataFrame(merged_rows).reset_index(drop=True)

    # ====== CLEAN NUMERIC CORE ======
    def clean_core(v):
        s = re.sub(r"[^0-9]", "", str(v or ""))
        return s

    erp_use["__core"] = erp_use["invoice_erp"].apply(clean_core)
    ven_use["__core"] = ven_use["invoice_ven"].apply(clean_core)

    # ====== REMOVE CANCELLED ======
    def remove_cancellations(df):
        cleaned = []
        grouped = df.groupby("invoice_erp" if "invoice_erp" in df.columns else "invoice_ven", dropna=False)
        for inv, g in grouped:
            if g.empty:
                continue
            amounts = g["__amt"].round(2).tolist()
            has_cancel_pair = any(a == -b for a in amounts for b in amounts if a != 0)
            if has_cancel_pair:
                g = g[~g["__amt"].isin([-x for x in amounts])]
            if not g.empty:
                cleaned.append(g.iloc[-1])
        return pd.DataFrame(cleaned)

    erp_use = remove_cancellations(erp_use)
    ven_use = remove_cancellations(ven_use)

    # ====== MATCHING (3 RULES ONLY) ======
    for e_idx, e in erp_use.iterrows():
        e_inv = str(e["invoice_erp"]).strip()
        e_core = e["__core"]
        e_amt = round(float(e["__amt"]), 2)
        e_date = e.get("date_erp")

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue

            v_inv = str(v["invoice_ven"]).strip()
            v_core = v["__core"]
            v_amt = round(float(v["__amt"]), 2)
            v_date = v.get("date_ven")

            amt_close = abs(e_amt - v_amt) < 0.05

            # Rule 1: Full invoice match
            if e_inv == v_inv and amt_close:
                match_type = "Full"
            # Rule 2: Last 3-digit match
            elif len(e_core) >= 3 and len(v_core) >= 3 and e_core[-3:] == v_core[-3:] and amt_close:
                match_type = "Last3"
            # Rule 3: Prefixless numeric match
            elif e_core.lstrip("0") == v_core.lstrip("0") and amt_close:
                match_type = "Prefixless"
            else:
                continue

            used_vendor_rows.add(v_idx)
            diff = round(e_amt - v_amt, 2)
            status = "Match" if abs(diff) < 0.05 else "Difference"

            matched.append({
                "Date (ERP)": e_date,
                "Date (Vendor)": v_date,
                "ERP Invoice": e_inv,
                "Vendor Invoice": v_inv,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status,
                "MatchType": match_type
            })
            break


    # Normalize invoices in all DataFrames before filtering missing
    erp_use["invoice_erp"] = erp_use["invoice_erp"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    ven_use["invoice_ven"] = ven_use["invoice_ven"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    
    for m in matched:
        m["ERP Invoice"] = str(m["ERP Invoice"]).strip().replace(".0", "")
        m["Vendor Invoice"] = str(m["Vendor Invoice"]).strip().replace(".0", "")

    # ====== BUILD MISSING TABLES (SAFE + CORRECT LOGIC) ======
    matched_erp = {m["ERP Invoice"] for m in matched}
    matched_ven = {m["Vendor Invoice"] for m in matched}
    
    def safe_cols(df, possible_cols):
        """Return only columns that actually exist in df."""
        return [c for c in possible_cols if c in df.columns]
    
    erp_cols = safe_cols(erp_use, ["date_erp", "invoice_erp", "__amt"])
    ven_cols = safe_cols(ven_use, ["date_ven", "invoice_ven", "__amt"])
    
    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][ven_cols] if "invoice_ven" in ven_use.columns else pd.DataFrame()
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][erp_cols] if "invoice_erp" in erp_use.columns else pd.DataFrame()
    
    if not missing_in_erp.empty:
        missing_in_erp = missing_in_erp.rename(columns={"date_ven": "Date", "invoice_ven": "Invoice", "__amt": "Amount"})
    else:
        missing_in_erp = pd.DataFrame(columns=["Date", "Invoice", "Amount"])
    
    if not missing_in_vendor.empty:
        missing_in_vendor = missing_in_vendor.rename(columns={"date_erp": "Date", "invoice_erp": "Invoice", "__amt": "Amount"})
    else:
        missing_in_vendor = pd.DataFrame(columns=["Date", "Invoice", "Amount"])
    
    matched_df = pd.DataFrame(matched)
    return matched_df, missing_in_erp, missing_in_vendor


# ======================================
# STREAMLIT UI
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)

    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    if "cif_ven" not in ven_df.columns or "cif_erp" not in erp_df.columns:
        st.error("❌ Missing CIF/VAT columns.")
        st.stop()

    vendor_cifs = sorted({str(x).strip().upper() for x in ven_df["cif_ven"].dropna().unique() if str(x).strip()})
    selected_cif = vendor_cifs[0] if len(vendor_cifs) == 1 else st.selectbox("Select Vendor CIF:", vendor_cifs)

    erp_df = erp_df[erp_df["cif_erp"].astype(str).str.strip().str.upper() == selected_cif]
    ven_df = ven_df[ven_df["cif_ven"].astype(str).str.strip().str.upper() == selected_cif]

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    total_match = len(matched[matched["Status"] == "Match"]) if not matched.empty else 0
    total_diff = len(matched[matched["Status"] == "Difference"]) if not matched.empty else 0
    st.success(f"✅ Recon complete for CIF {selected_cif}: {total_match} matched, {total_diff} differences")

    def highlight_row(row):
        if row.get("Status") == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row.get("Status") == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        else:
            return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")

    st.subheader("❌ Missing in ERP (invoices found in vendor but not ERP)")
    if not erp_missing.empty:
        st.dataframe(
            erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True,
        )
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (invoices found in ERP but not vendor)")
    if not ven_missing.empty:
        st.dataframe(
            ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True,
        )
    else:
        st.success("✅ No missing invoices in Vendor file.")

    st.download_button("⬇️ Matched CSV", matched.to_csv(index=False).encode("utf-8"), "matched.csv", "text/csv")
    st.download_button("⬇️ Missing ERP CSV", erp_missing.to_csv(index=False).encode("utf-8"), "missing_erp.csv", "text/csv")
    st.download_button("⬇️ Missing Vendor CSV", ven_missing.to_csv(index=False).encode("utf-8"), "missing_vendor.csv", "text/csv")
else:
    st.info("Please upload both ERP and Vendor files to begin.")
