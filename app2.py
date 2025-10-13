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
    "invoice": ["factura", "fact", "fac", "doc", "documento", "invoice", "παραστατικό"],
    "credit_note": ["abono", "nota de crédito", "nota credito", "nc", "avoir", "πιστωτικό"],
    "payment": ["pago", "transferencia", "remesa", "domiciliación", "cobro", "recepción", "payment", "κατάθεση"],
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
    """Map multilingual headers to unified names (robust, with diagnostics)."""
    mapping = {
        "invoice": [
            "alternative document", "alt document", "invoice", "factura",
            "nº factura", "no", "nro", "num", "numero", "document", "doc", "παραστατικό"
        ],
        "credit": ["credit", "haber", "crédito", "credito"],
        "debit": [
            "debit", "debe", "cargo", "importe", "valor", "total", "totale",
            "amount", "importe total", "document value", "charge", "charges"
        ],
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

    # Fallback creation if still missing key columns
    required = ["debit", "credit"]
    for coltype in required:
        cname = f"{coltype}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0

    # ✅ Debug info: show what was mapped
    with st.expander(f"🧩 Column mapping detected for {tag.upper()} file"):
        st.write({k: v for k, v in rename_map.items()})
        st.write("✅ Columns after normalization:", list(out.columns))

    return out


def extract_core_invoice(inv):
    """Extract a meaningful invoice tail key."""
    if isinstance(inv, (pd.Series, list)):
        inv = inv.iloc[0] if isinstance(inv, pd.Series) else inv[0]
    if inv is None or (isinstance(inv, float) and pd.isna(inv)):
        return ""
    s = re.sub(r"[^A-Za-z0-9]", "", str(inv).upper())
    m = re.search(r"([A-Z]*\d{2,6})$", s)
    return m.group(1) if m else (s[-4:] if len(s) > 4 else s)


# ============================================================
# ✅ FINAL VERSION — Robust 3+ Digit Invoice Matcher
# ============================================================
def share_3plus_digits(a, b):
    """
    Determine if two invoice references share any 3+ consecutive digits,
    ignoring prefixes, letters, and special characters (e.g. PF, DS, /, --, etc.).
    Never matches between Invoice and Credit Note (INV vs CN).
    """
    a = str(a).upper().strip()
    b = str(b).upper().strip()

    # Skip matching between Invoice and Credit Note types
    if (a.startswith("CN") and not b.startswith("CN")) or (b.startswith("CN") and not a.startswith("CN")):
        return False

    # Keep only digits
    a_digits = re.sub(r"\D", "", a)
    b_digits = re.sub(r"\D", "", b)
    if not a_digits or not b_digits:
        return False

    # Check all possible overlapping numeric substrings of length ≥3
    for i in range(len(a_digits) - 2):
        for length in range(6, 2, -1):  # 6,5,4,3
            seq = a_digits[i:i + length]
            if len(seq) >= 3 and seq in b_digits:
                return True
    return False


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched, matched_erp, matched_ven = [], set(), set()

    req_erp = ["invoice_erp", "credit_erp", "debit_erp", "reason_erp", "date_erp", "cif_erp"]
    req_ven = ["invoice_ven", "debit_ven", "date_ven", "cif_ven"]
    for c in req_erp:
        if c not in erp_df.columns:
            st.error(f"❌ ERP file missing column: {c}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    for c in req_ven:
        if c not in ven_df.columns:
            st.error(f"❌ Vendor file missing column: {c}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    erp_df["__core"] = erp_df["invoice_erp"].astype(str).apply(extract_core_invoice)
    ven_df["__core"] = ven_df["invoice_ven"].astype(str).apply(extract_core_invoice)

    erp_df["__doctype"] = erp_df.apply(lambda r: classify_erp_doc(r.get("reason_erp"), r.get("credit_erp"), r.get("debit_erp")), axis=1)
    erp_df["__amt"] = erp_df.apply(
        lambda r: normalize_number(r["credit_erp"]) if r["__doctype"] == "INV"
                  else (-normalize_number(r["debit_erp"]) if r["__doctype"] == "CN" else 0.0),
        axis=1
    )
    ven_df["__amt"] = ven_df.apply(lambda r: normalize_number(r.get("debit_ven", 0.0)), axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df.copy()

    for _, e_row in erp_use.iterrows():
        e_inv = str(e_row["invoice_erp"]).strip()
        e_core = e_row["__core"]
        e_amt = round(float(e_row["__amt"]), 2)
        e_date = e_row.get("date_erp")

        best_v = None
        best_score = -1

        for _, v_row in ven_use.iterrows():
            v_inv = str(v_row["invoice_ven"]).strip()
            v_amt = round(float(v_row["__amt"]), 2)
            v_date = v_row.get("date_ven")

            if v_inv in matched_ven:
                continue

            # 3-digit overlap rule (the key update)
            digits3 = share_3plus_digits(e_inv, v_inv)
            fuzzy = fuzz.ratio(e_inv, v_inv) if e_inv and v_inv else 0

            if not digits3 and fuzzy < 90:
                continue

            amt_close = abs(e_amt - v_amt) < 0.05
            score = (70 if digits3 else 0) + fuzzy + (60 if amt_close else 0)
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

    erp_use["__clean_inv"] = erp_use["invoice_erp"].str.replace(r"[^A-Z0-9]", "", regex=True)
    ven_use["__clean_inv"] = ven_use["invoice_ven"].str.replace(r"[^A-Z0-9]", "", regex=True)
    matched_erp_clean = {i.upper() for i in matched_erp}
    matched_ven_clean = {i.upper() for i in matched_ven}

    erp_missing = (
        erp_use[~erp_use["__clean_inv"].isin(matched_erp_clean)]
        .loc[:, ["date_erp", "invoice_erp", "__amt"]]
        .rename(columns={"date_erp": "Date", "invoice_erp": "Invoice", "__amt": "Amount"})
        .reset_index(drop=True)
    )

    ven_missing = (
        ven_use[~ven_use["__clean_inv"].isin(matched_ven_clean)]
        .loc[:, ["date_ven", "invoice_ven", "__amt"]]
        .rename(columns={"date_ven": "Date", "invoice_ven": "Invoice", "__amt": "Amount"})
        .reset_index(drop=True)
    )

    df_matched = pd.DataFrame(matched)
    return df_matched, erp_missing, ven_missing


# ======================================
# STREAMLIT UI
# ======================================
st.write("Upload your ERP Export and Vendor Statement to reconcile invoices automatically.")

uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp)
    ven_raw = pd.read_excel(uploaded_vendor)

    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    if "cif_ven" not in ven_df.columns or "cif_erp" not in erp_df.columns:
        st.error("❌ Missing CIF/VAT columns.")
        st.stop()

    vendor_cifs = sorted({str(x).strip().upper() for x in ven_df["cif_ven"].dropna().unique() if str(x).strip()})
    selected_cif = vendor_cifs[0] if len(vendor_cifs) == 1 else st.selectbox("Select Vendor CIF:", vendor_cifs)

    erp_df = erp_df[erp_df["cif_erp"].astype(str).str.strip().str.upper() == selected_cif].copy()
    ven_df = ven_df[ven_df["cif_ven"].astype(str).str.strip().str.upper() == selected_cif].copy()

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    total_match = len(matched[matched["Status"] == "Match"]) if not matched.empty else 0
    total_diff = len(matched[matched["Status"] == "Difference"]) if not matched.empty else 0
    total_missing = len(erp_missing) + len(ven_missing)
    st.success(f"✅ Recon complete for CIF {selected_cif}: {total_match} matched, {total_diff} differences, {total_missing} missing")

    def highlight_row(row):
        if row.get("Status") == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row.get("Status") == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        else:
            return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)

    st.subheader("❌ Missing in ERP")
    st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)

    st.subheader("❌ Missing in Vendor")
    st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)

    st.download_button("⬇️ Matched CSV", matched.to_csv(index=False).encode("utf-8"), "matched.csv", "text/csv")
    st.download_button("⬇️ Missing in ERP CSV", erp_missing.to_csv(index=False).encode("utf-8"), "missing_erp.csv", "text/csv")
    st.download_button("⬇️ Missing in Vendor CSV", ven_missing.to_csv(index=False).encode("utf-8"), "missing_vendor.csv", "text/csv")
else:
    st.info("Please upload both ERP and Vendor Excel files to begin.")
