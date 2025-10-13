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
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": ["invoice", "factura", "document", "doc", "nº", "num"],
        "credit": ["credit", "haber", "credito"],
        "debit": ["debit", "debe", "cargo", "importe", "valor", "amount", "document value", "charge"],
        "reason": ["reason", "motivo", "concepto", "descripcion"],
        "cif": ["cif", "nif", "vat", "tax"],
        "date": ["date", "fecha", "data"],
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


_digit_seq = re.compile(r"\d{3,}")  # 3+ digit sequence matcher

def share_3plus_digits(a, b):
    """True if a and b share any continuous 3+ digit substring."""
    A = set(m.group(0) for m in _digit_seq.finditer(str(a)))
    if not A:
        return False
    for m in _digit_seq.finditer(str(b)):
        if m.group(0) in A:
            return True
    return False


def clean_core(v):
    """Extract numeric core (last 3–6 digits) ignoring special characters."""
    s = re.sub(r"[^0-9]", "", str(v))
    if len(s) > 6:
        return s[-6:]
    return s


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    matched_pairs = set()

    # ============== PREP ERP ==================
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

    # ============== PREP VENDOR ==================
    ven_df["__doctype"] = ven_df.apply(
        lambda r: "CN" if normalize_number(r.get("debit_ven")) < 0 else "INV",
        axis=1
    )
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven"))), axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # ============== CLEAN CORE (numeric) ==================
    def clean_core(v):
        s = re.sub(r"[^0-9]", "", str(v or ""))
        return s[-6:] if len(s) >= 6 else s

    erp_use["__core"] = erp_use["invoice_erp"].apply(clean_core)
    ven_use["__core"] = ven_use["invoice_ven"].apply(clean_core)

    # ============== MAIN MATCHING LOOP ==================
    for _, e in erp_use.iterrows():
        e_inv = str(e["invoice_erp"]).strip()
        e_core = clean_core(e_inv)
        e_amt = round(float(e["__amt"]), 2)
        e_date = e.get("date_erp")

        best = None
        best_score = -1

        for _, v in ven_use.iterrows():
            v_inv = str(v["invoice_ven"]).strip()
            v_core = clean_core(v_inv)
            v_amt = round(float(v["__amt"]), 2)
            v_date = v.get("date_ven")

            # avoid reusing exact same core pair
            if (e_core, v_core) in matched_pairs:
                continue

            # ---------- 3-digit rule ----------
            # if any sequence of 3+ digits overlaps between ERP and vendor core, it's a match
            digits_erp = re.findall(r"\d{3,}", e_core)
            digits_ven = re.findall(r"\d{3,}", v_core)
            three_match = any(d in v_core for d in digits_erp if len(d) >= 3) or any(d in e_core for d in digits_ven if len(d) >= 3)

            fuzzy = fuzz.ratio(e_inv, v_inv)
            amt_close = abs(e_amt - v_amt) < 0.05

            if three_match or fuzzy > 90 or e_core == v_core:
                score = fuzzy + (100 if three_match or e_core == v_core else 0) + (60 if amt_close else 0)
                if score > best_score:
                    best_score = score
                    best = (v_inv, v_core, v_amt, v_date)

        if best:
            v_inv, v_core, v_amt, v_date = best
            diff = round(e_amt - v_amt, 2)
            status = "Match" if abs(diff) < 0.05 else "Difference"

            matched.append({
                "Date (ERP)": e_date,
                "Date (Vendor)": v_date,
                "ERP Invoice": e_inv if e_inv else "(inferred)",
                "Vendor Invoice": v_inv,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status
            })

            matched_pairs.add((e_core, v_core))

    # ============== BUILD MISSING TABLES ==================
    erp_matched_cores = {pair[0] for pair in matched_pairs}
    ven_matched_cores = {pair[1] for pair in matched_pairs}

    def core_of(v): return clean_core(v)

    # invoices present in vendor but not ERP → Missing in ERP
    ven_missing = (
        ven_use[~ven_use["__core"].isin(erp_matched_cores)]
        .loc[:, ["date_ven", "invoice_ven", "__amt"]]
        .rename(columns={"date_ven": "Date", "invoice_ven": "Invoice", "__amt": "Amount"})
        .reset_index(drop=True)
    )

    # invoices present in ERP but not vendor → Missing in Vendor
    erp_missing = (
        erp_use[~erp_use["__core"].isin(ven_matched_cores)]
        .loc[:, ["date_erp", "invoice_erp", "__amt"]]
        .rename(columns={"date_erp": "Date", "invoice_erp": "Invoice", "__amt": "Amount"})
        .reset_index(drop=True)
    )

    df_matched = pd.DataFrame(matched)
    return df_matched, erp_missing, ven_missing
    # --- define clean numeric core ---
    def clean_core(v):
        """Return last 4–6 digits of any invoice for core match."""
        s = re.sub(r"[^0-9]", "", str(v))
        return s[-6:] if len(s) >= 6 else s

    # --- main matching loop ---
    for _, e_row in erp_use.iterrows():
        e_inv = str(e_row["invoice_erp"]).strip()
        e_amt = round(float(e_row["__amt"]), 2)
        e_date = e_row.get("date_erp")
        e_core = clean_core(e_inv)

        best_v = None
        best_score = -1

        for _, v_row in ven_use.iterrows():
            v_inv = str(v_row["invoice_ven"]).strip()
            v_core = clean_core(v_inv)
            v_amt = round(float(v_row["__amt"]), 2)
            v_date = v_row.get("date_ven")

            # Skip if vendor core already used by the exact same ERP core
            pair_key = f"{e_core}-{v_core}"
            if pair_key in matched_ven_core:
                continue

            # --- 3-digit rule ---
            three_match = False
            for match in re.findall(r"\d{3,}", v_core):
                if match in e_core or e_core.endswith(match):
                    three_match = True
                    break

            fuzzy = fuzz.ratio(e_inv, v_inv)
            amt_close = abs(e_amt - v_amt) < 0.05
            score = fuzzy + (60 if amt_close else 0) + (100 if three_match or e_core == v_core else 0)

            if score > best_score:
                best_score = score
                best_v = (v_inv, v_core, v_amt, v_date)

        if best_v:
            v_inv, v_core, v_amt, v_date = best_v
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
            matched_ven_core.add(f"{e_core}-{v_core}")  # track pair instead of single vendor invoice

    # --- missing logic ---
    def clean_invoice(v):
        return re.sub(r"[^A-Z0-9]", "", str(v).upper().strip())

    erp_use["__clean_inv"] = erp_use["invoice_erp"].apply(clean_invoice)
    ven_use["__clean_inv"] = ven_use["invoice_ven"].apply(clean_invoice)

    matched_erp_clean = {clean_invoice(i) for i in matched_erp}
    matched_ven_clean = {re.sub(r"[^0-9]", "", i[-6:]) for i in matched_ven_core}

    erp_missing = (
        erp_use[~erp_use["__clean_inv"].isin(matched_erp_clean)]
        .loc[:, ["date_erp", "invoice_erp", "__amt"]]
        .rename(columns={"date_erp": "Date", "invoice_erp": "Invoice", "__amt": "Amount"})
        .reset_index(drop=True)
    )

    ven_missing = (
        ven_use[~ven_use["invoice_ven"].apply(lambda x: clean_core(x)).isin(matched_ven_clean)]
        .loc[:, ["date_ven", "invoice_ven", "__amt"]]
        .rename(columns={"date_ven": "Date", "invoice_ven": "Invoice", "__amt": "Amount"})
        .reset_index(drop=True)
    )

    df_matched = pd.DataFrame(matched)
    return df_matched, erp_missing, ven_missing

    # --- Missing logic ---
    def clean_invoice(v):
        return re.sub(r"[^A-Z0-9]", "", str(v).upper().strip())

    erp_use["__clean_inv"] = erp_use["invoice_erp"].apply(clean_invoice)
    ven_use["__clean_inv"] = ven_use["invoice_ven"].apply(clean_invoice)

    matched_erp_clean = {clean_invoice(i) for i in matched_erp}
    matched_ven_clean = {clean_invoice(i) for i in matched_ven_core}

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
st.write("Upload your ERP Export (Credit = Invoice / Charge = Credit Note) and Vendor Statement (Document Value).")

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
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
                     use_container_width=True)
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (invoices found in ERP but not vendor)")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
                     use_container_width=True)
    else:
        st.success("✅ No missing invoices in Vendor file.")

    st.download_button("⬇️ Matched CSV", matched.to_csv(index=False).encode("utf-8"), "matched.csv", "text/csv")
    st.download_button("⬇️ Missing ERP CSV", erp_missing.to_csv(index=False).encode("utf-8"), "missing_erp.csv", "text/csv")
    st.download_button("⬇️ Missing Vendor CSV", ven_missing.to_csv(index=False).encode("utf-8"), "missing_vendor.csv", "text/csv")
else:
    st.info("Please upload both ERP and Vendor files to begin.")
