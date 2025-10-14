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
    """Map multilingual headers to unified names — fully optimized for Spanish vendor statements."""
    mapping = {
        # 🔢 Invoice / Document number detection
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura"
        ],

        # 💳 Credit note / Abono / Haber detection
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber"
        ],

        # 💰 Debit / Document Value / Total detection
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge",
            "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto"
        ],

        # 🗒️ Reason / Description detection
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "descriptivo", "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion"
        ],

        # 🧾 CIF / VAT / NIF detection
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal"
        ],

        # 📅 Date detection
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
    ven_df["__doctype"] = ven_df.apply(
        lambda r: "CN" if normalize_number(r.get("debit_ven")) < 0 else "INV",
        axis=1
    )
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven"))), axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # ====== CLEAN NUMERIC CORE ======
    def clean_core(v):
        s = re.sub(r"[^0-9]", "", str(v or ""))
        return s[-6:] if len(s) >= 6 else s

    erp_use["__core"] = erp_use["invoice_erp"].apply(clean_core)
    ven_use["__core"] = ven_use["invoice_ven"].apply(clean_core)

    # ====== MATCHING (pick best vendor per ERP) ======
    for e_idx, e in erp_use.iterrows():
        e_inv = str(e["invoice_erp"]).strip()
        e_core = e["__core"]
        e_amt = round(float(e["__amt"]), 2)
        e_date = e.get("date_erp")

        best_score = -1
        best_v = None

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue

            v_inv = str(v["invoice_ven"]).strip()
            v_core = v["__core"]
            v_amt = round(float(v["__amt"]), 2)
            v_date = v.get("date_ven")

            # === numeric overlap rule ===
            digits_erp = re.findall(r"\d{3,}", e_core)
            digits_ven = re.findall(r"\d{3,}", v_core)
            three_match = any(d in v_core or v_core.endswith(d) or e_core.endswith(d) for d in digits_erp if len(d) >= 3)

            fuzzy = fuzz.ratio(e_inv, v_inv)
            amt_close = abs(e_amt - v_amt) < 0.05

            score = fuzzy + (80 if three_match or e_core == v_core else 0) + (50 if amt_close else 0)

            if score > best_score:
                best_score = score
                best_v = (v_idx, v_inv, v_core, v_amt, v_date)

        if best_v and best_score >= 120:  # threshold for confident match
            v_idx, v_inv, v_core, v_amt, v_date = best_v
            used_vendor_rows.add(v_idx)
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

    # ====== BUILD MISSING TABLES ======
    # ====== BUILD MISSING TABLES (3+ digit overlap aware) ======

        # ====== BUILD MISSING TABLES (final fix) ======

       # ====== BUILD MISSING TABLES (final symmetric logic) ======

        # ====== BUILD MISSING TABLES (final clean symmetric fix) ======

    def extract_tokens(s: str):
        """Extract all 3+ digit sequences from an invoice string."""
        return set(re.findall(r"\d{3,}", str(s or "")))

    # Build token sets for ERP and Vendor
    erp_tokens = {str(e): extract_tokens(e) for e in erp_use["__core"]}
    ven_tokens = {str(v): extract_tokens(v) for v in ven_use["__core"]}

    matched_erp_invs = {m["ERP Invoice"] for m in matched}
    matched_ven_invs = {m["Vendor Invoice"] for m in matched}

    # Combine token pools
    all_erp_tokens = set().union(*erp_tokens.values()) if len(erp_tokens) else set()
    all_ven_tokens = set().union(*ven_tokens.values()) if len(ven_tokens) else set()

    # --- Missing in ERP (vendor invoices with no shared 3+ digits in ANY ERP invoice) ---
    ven_missing_list = []
    for _, row in ven_use.iterrows():
        inv = str(row["invoice_ven"])
        core = str(row["__core"])
        if inv in matched_ven_invs:
            continue
        # Missing if no 3-digit overlap with any ERP core
        if len(extract_tokens(core) & all_erp_tokens) == 0:
            ven_missing_list.append(row)

    # --- Missing in Vendor (ERP invoices with no shared 3+ digits in ANY vendor invoice) ---
    # 👇 We merge this logic into Missing in ERP instead
    erp_missing_list = []
    for _, row in erp_use.iterrows():
        inv = str(row["invoice_erp"])
        core = str(row["__core"])
        if inv in matched_erp_invs:
            continue
        if len(extract_tokens(core) & all_ven_tokens) == 0:
            # Move all these ERP-only invoices into the same Missing in ERP table
            ven_missing_list.append(row)

    # --- Create clean DataFrames ---
    # --- Create clean DataFrames (remove NaN/None rows) ---
    if ven_missing_list:
        combined_rows = []
        for r in ven_missing_list:
            rec = {}
            # unify column names whether it came from ERP or Vendor
            rec["Date"] = r.get("date_ven") or r.get("date_erp")
            rec["Invoice"] = r.get("invoice_ven") or r.get("invoice_erp")
            rec["Amount"] = r.get("__amt")
            # skip completely empty or None rows
            if rec["Invoice"] and str(rec["Invoice"]).lower() != "none":
                combined_rows.append(rec)
        missing_erp_final = pd.DataFrame(combined_rows)
    else:
        missing_erp_final = pd.DataFrame(columns=["Date", "Invoice", "Amount"])

    # Remove any accidental blank or duplicate rows
    if not missing_erp_final.empty:
        missing_erp_final = (
            missing_erp_final.dropna(subset=["Invoice"])
            .query("Invoice != 'None'")
            .drop_duplicates(subset=["Invoice"])
            .reset_index(drop=True)
        )

    # Missing in Vendor now always empty
    missing_vendor_final = pd.DataFrame(columns=["Date", "Invoice", "Amount"])
 # ✅ Convert Amount column to numeric (if text)
    for df in [matched, missing_erp_final, missing_vendor_final]:
        if "Amount" in df.columns:
            df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)
    return pd.DataFrame(matched), missing_erp_final, missing_vendor_final
# ======================================
# STREAMLIT UI
# ======================================

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
