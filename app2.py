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
    if isinstance(v, (pd.Series, list)):
        v = v.iloc[0] if isinstance(v, pd.Series) else v[0]
    if v is None:
        return 0.0
    try:
        if isinstance(v, float) and pd.isna(v):
            return 0.0
    except Exception:
        pass
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


SPANISH_REASON_MAP = {
    "invoice": [
        "factura", "fact", "fac", "doc", "documento", "invoice", "παραστατικό"
    ],
    "credit_note": [
        "abono", "nota de crédito", "nota credito", "nc", "avoir", "πιστωτικό"
    ],
    "payment": [
        "pago", "transferencia", "remesa", "domiciliación", "cobro", "recepción",
        "payment", "κατάθεση"
    ],
}

def classify_erp_doc(reason_text, credit, charge):
    """Return INV / CN / PAYMENT / UNKNOWN using Reason + amount columns."""
    r = (str(reason_text) if reason_text is not None else "").strip().lower()
    def has_any(words): 
        return any(w in r for w in words)
    if has_any(SPANISH_REASON_MAP["credit_note"]):
        return "CN"
    if has_any(SPANISH_REASON_MAP["invoice"]):
        return "INV"
    if has_any(SPANISH_REASON_MAP["payment"]):
        return "PAYMENT"
    # Fallback by amounts
    c = normalize_number(credit)
    ch = normalize_number(charge)
    if c > 0 and ch == 0:
        return "INV"
    if ch > 0 and c == 0:
        return "PAYMENT"
    return "UNKNOWN"


def normalize_columns(df, tag):
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "alternative document", "alt document", "invoice", "factura",
            "nº factura", "no", "nro", "num", "numero", "document", "doc", "παραστατικό"
        ],
        "credit": ["credit", "haber", "crédito", "credito"],
        "debit": ["debit", "debe", "cargo", "importe", "valor", "total", "totale", "amount", "importe total"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "glosa", "observaciones"],
        "cif": ["cif", "nif", "vat", "tax", "vat number", "tax id", "afm", "vat id", "num. identificacion fiscal"],
        "date": ["date", "fecha", "fech", "ημερομηνία", "data"],
    }
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    for k, vals in mapping.items():
        for col, low in cols_lower.items():
            if any(v in low for v in vals):
                rename_map[col] = f"{k}_{tag}"
    out = df.rename(columns=rename_map)
    return out


def extract_core_invoice(inv):
    """
    Extract a meaningful invoice tail key:
    - keep alnum uppercase
    - prefer last A* + 2-6 digits (e.g., PF006, AB1234)
    - fallback last 4 chars
    """
    if isinstance(inv, (pd.Series, list)):
        inv = inv.iloc[0] if isinstance(inv, pd.Series) else inv[0]
    if inv is None or (isinstance(inv, float) and pd.isna(inv)):
        return ""
    s = re.sub(r"[^A-Za-z0-9]", "", str(inv).upper())
    m = re.search(r"([A-Z]*\d{2,6})$", s)
    return m.group(1) if m else (s[-4:] if len(s) > 4 else s)

_digit_seq = re.compile(r"\d{3,}")  # any 3+ continuous digits

def share_3plus_digits(a, b):
    """True if a and b share any continuous 3+ digit substring."""
    A = set(m.group(0) for m in _digit_seq.finditer(a))
    if not A:
        return False
    for m in _digit_seq.finditer(b):
        if m.group(0) in A:
            return True
    return False


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched, matched_erp, matched_ven = [], set(), set()

    # Required columns check
    req_erp = ["invoice_erp", "credit_erp", "debit_erp", "reason_erp", "date_erp", "cif_erp"]
    req_ven = ["invoice_ven", "debit_ven", "credit_ven", "date_ven", "cif_ven"]
    for c in req_erp:
        if c not in erp_df.columns:
            st.error(f"❌ ERP file missing column: {c}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    for c in req_ven:
        if c not in ven_df.columns:
            st.error(f"❌ Vendor file missing column: {c}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Add core keys
    erp_df["__core"] = erp_df["invoice_erp"].astype(str).apply(extract_core_invoice)
    ven_df["__core"] = ven_df["invoice_ven"].astype(str).apply(extract_core_invoice)

    # Pre-compute doc types + effective amounts
    # ERP amount: INV -> +Credit ; CN -> -Charge ; PAYMENT -> ignored
    erp_df["__doctype"] = erp_df.apply(
        lambda r: classify_erp_doc(r.get("reason_erp"), r.get("credit_erp"), r.get("debit_erp")),
        axis=1
    )
    erp_df["__amt"] = erp_df.apply(
        lambda r: normalize_number(r["credit_erp"]) if r["__doctype"] == "INV"
                  else (-normalize_number(r["debit_erp"]) if r["__doctype"] == "CN" else 0.0),
        axis=1
    )

    # Vendor amount: prefer debit_ven (Factura). Ignore credit_ven as payments (unless later extend to vendor CN)
    # If only total without structure, normalize_columns already mapped to debit_ven via synonyms
    ven_df["__amt"] = ven_df.apply(
        lambda r: normalize_number(r.get("debit_ven", 0.0)),
        axis=1
    )

    # Filter: consider only INV and CN on ERP side
    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df.copy()

    # Matching loop
    for _, e_row in erp_use.iterrows():
        e_inv = str(e_row["invoice_erp"]).strip()
        e_core = e_row["__core"]
        e_amt = round(float(e_row["__amt"]), 2)
        e_date = e_row.get("date_erp")
        e_key = f"{e_core}|{e_amt}"

        best_v = None
        best_score = -1

        for _, v_row in ven_use.iterrows():
            v_inv = str(v_row["invoice_ven"]).strip()
            v_core = v_row["__core"]
            v_amt = round(float(v_row["__amt"]), 2)
            v_date = v_row.get("date_ven")

            # Skip if already matched this vendor invoice
            if v_inv in matched_ven:
                continue

            # Heuristics
            exact = (e_inv == v_inv)
            core_eq = (e_core == v_core) and e_core != ""
            ends = (e_core.endswith(v_core) or v_core.endswith(e_core)) and e_core != "" and v_core != ""
            digits3 = share_3plus_digits(e_inv.upper(), v_inv.upper())
            fuzzy = fuzz.ratio(e_inv, v_inv) if e_inv and v_inv else 0

            rule_hit = exact or core_eq or ends or digits3 or (fuzzy > 90)
            if not rule_hit:
                continue

            # Prefer same amount if multiple candidates
            amt_close = abs(e_amt - v_amt) < 0.05
            score = (100 if exact else 0) + (40 if core_eq else 0) + (30 if ends else 0) + (35 if digits3 else 0) + fuzzy + (60 if amt_close else 0)
            if score > best_score:
                best_score = score
                best_v = (v_inv, v_amt, v_date)

        if best_v:
            v_inv, v_amt, v_date = best_v
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
                "Status": status
            })
            matched_erp.add(e_inv)
            matched_ven.add(v_inv)

    # Clean keys for exclusion in missing tables
    def clean_invoice(v):
        return re.sub(r"[^A-Z0-9]", "", str(v).strip().upper())
    erp_use["__clean_inv"] = erp_use["invoice_erp"].apply(clean_invoice)
    ven_use["__clean_inv"] = ven_use["invoice_ven"].apply(clean_invoice)
    matched_erp_clean = {clean_invoice(i) for i in matched_erp}
    matched_ven_clean = {clean_invoice(i) for i in matched_ven}

    # Build missing tables (only unmatched)
    erp_missing = (
        erp_use[~erp_use["__clean_inv"].isin(matched_erp_clean)]
        .loc[:, ["date_erp", "invoice_erp", "__amt"]]
        .rename(columns={
            "date_erp": "Date",
            "invoice_erp": "Invoice",
            "__amt": "Amount"
        })
        .reset_index(drop=True)
    )

    ven_missing = (
        ven_use[~ven_use["__clean_inv"].isin(matched_ven_clean)]
        .loc[:, ["date_ven", "invoice_ven", "__amt"]]
        .rename(columns={
            "date_ven": "Date",
            "invoice_ven": "Invoice",
            "__amt": "Amount"
        })
        .reset_index(drop=True)
    )

    df_matched = pd.DataFrame(matched)
    return df_matched, erp_missing, ven_missing


# ======================================
# STREAMLIT UI
# ======================================
st.write("Upload your ERP Export (Date / Alternative Document / Reason / Charge / Credit / CIF) and Vendor Statement (Factura/No/Doc + Debe/Importe/Total + CIF).")

uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp)
    ven_raw = pd.read_excel(uploaded_vendor)

    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    # Ensure minimal columns exist for selection
    if "cif_ven" not in ven_df.columns:
        st.error("❌ Vendor file must contain a CIF/NIF/VAT column.")
        st.stop()
    if "cif_erp" not in erp_df.columns:
        st.error("❌ ERP file must contain a CIF/NIF/VAT column.")
        st.stop()

    # Pick vendor CIF from Vendor file (one vendor at a time)
    vendor_cifs = sorted({str(x).strip().upper() for x in ven_df["cif_ven"].dropna().unique() if str(x).strip()})
    selected_cif = vendor_cifs[0] if len(vendor_cifs) == 1 else st.selectbox("Select Vendor CIF to reconcile:", vendor_cifs)

    # Filter both sides by CIF
    erp_df = erp_df[erp_df["cif_erp"].astype(str).str.strip().str.upper() == selected_cif].copy()
    ven_df = ven_df[ven_df["cif_ven"].astype(str).str.strip().str.upper() == selected_cif].copy()

    # Basic column presence for matching
    for needed in ["invoice_erp", "credit_erp", "debit_erp", "reason_erp", "date_erp"]:
        if needed not in erp_df.columns:
            st.error(f"❌ ERP missing column: {needed}")
            st.stop()
    for needed in ["invoice_ven", "debit_ven", "date_ven"]:
        if needed not in ven_df.columns:
            # If vendor has only a generic total/amount column it should have been mapped to debit_ven already via synonyms
            st.error(f"❌ Vendor missing column: {needed}")
            st.stop()

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    # --- Summary ---
    total_match = len(matched[matched["Status"] == "Match"]) if not matched.empty else 0
    total_diff = len(matched[matched["Status"] == "Difference"]) if not matched.empty else 0
    total_missing = len(erp_missing) + len(ven_missing)
    st.success(f"✅ Recon complete for CIF {selected_cif}: {total_match} matched, {total_diff} differences, {total_missing} missing")

    # --- Styling helpers ---
    def highlight_row(row):
        if row.get("Status") == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row.get("Status") == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        else:
            return [''] * len(row)

    # --- Display tables ---
    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches or differences found.")

    st.subheader("❌ Missing in ERP (for selected CIF)")
    if not erp_missing.empty:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in ERP for this vendor.")

    st.subheader("❌ Missing in Vendor (for selected CIF)")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in Vendor file for this vendor.")

    # --- Downloads ---
    st.download_button("⬇️ Matched/Differences CSV", matched.to_csv(index=False).encode("utf-8"), "matched_results.csv", "text/csv")
    st.download_button("⬇️ Missing in ERP CSV", erp_missing.to_csv(index=False).encode("utf-8"), "missing_in_erp.csv", "text/csv")
    st.download_button("⬇️ Missing in Vendor CSV", ven_missing.to_csv(index=False).encode("utf-8"), "missing_in_vendor.csv", "text/csv")
else:
    st.info("Please upload both ERP and Vendor files to begin.")
