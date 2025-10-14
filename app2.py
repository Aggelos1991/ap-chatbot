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
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal"
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
    ven_df["__doctype"] = ven_df.apply(
        lambda r: "CN" if normalize_number(r.get("debit_ven")) < 0 else "INV",
        axis=1
    )
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven"))), axis=1)

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
            net = round(total_inv + total_cn, 2)  # CN amounts are negative in ERP

            # Keep one line with net amount
            base_row = inv_rows.iloc[0].copy()
            base_row["__amt"] = net
            merged_rows.append(base_row)
        else:
            # If only invoices or only CNs exist, keep all
            for _, row in group.iterrows():
                merged_rows.append(row)

    erp_use = pd.DataFrame(merged_rows).reset_index(drop=True)


    # ====== CLEAN NUMERIC CORE ======
    def clean_core(v):
        s = re.sub(r"[^0-9]", "", str(v or ""))
        return s[-6:] if len(s) >= 6 else s

    erp_use["__core"] = erp_use["invoice_erp"].apply(clean_core)
    ven_use["__core"] = ven_use["invoice_ven"].apply(clean_core)
        # ====== REMOVE CANCELLED OR DUPLICATE INVOICES ======
    def remove_cancellations(df):
        """Remove invoices where positive & negative entries cancel out, keeping only final valid ones."""
        cleaned = []
        grouped = df.groupby("invoice_erp" if "invoice_erp" in df.columns else "invoice_ven", dropna=False)
        for inv, g in grouped:
            if g.empty:
                continue
            # If both positive and negative of same absolute value exist → skip both
            amounts = g["__amt"].round(2).tolist()
            has_cancel_pair = any(a == -b for a in amounts for b in amounts if a != 0)
            if has_cancel_pair:
                # keep only entries not fully cancelled (e.g. small residuals or new final invoice)
                g = g[~g["__amt"].isin([-x for x in amounts])]
            # keep the last (usually latest) row if still duplicates
            if not g.empty:
                cleaned.append(g.iloc[-1])
        return pd.DataFrame(cleaned)

    erp_use = remove_cancellations(erp_use)
    ven_use = remove_cancellations(ven_use)


   def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    # ===== Clean numeric cores =====
    def clean_core(v):
        return re.sub(r"[^0-9]", "", str(v or ""))

    erp_df["__core"] = erp_df["invoice_erp"].apply(clean_core)
    ven_df["__core"] = ven_df["invoice_ven"].apply(clean_core)

    # ===== Match loop =====
    for e_idx, e in erp_df.iterrows():
        e_inv = str(e["invoice_erp"]).strip()
        e_core = e["__core"]
        e_amt = normalize_number(e.get("credit_erp")) - normalize_number(e.get("debit_erp"))
        e_date = e.get("date_erp")

        best_v = None
        best_score = -1

        for v_idx, v in ven_df.iterrows():
            if v_idx in used_vendor_rows:
                continue

            v_inv = str(v["invoice_ven"]).strip()
            v_core = v["__core"]
            v_amt = normalize_number(v.get("debit_ven")) - normalize_number(v.get("credit_ven"))
            v_date = v.get("date_ven")

            # === Rule 1: Exact full invoice match ===
            exact_match = e_inv.lower() == v_inv.lower()

            # === Rule 2: Last 3 digits match ===
            last3_match = False
            if len(e_core) >= 3 and len(v_core) >= 3:
                last3_match = e_core[-3:] == v_core[-3:]

            # === Rule 3: Prefix/suffix short numeric match (PSF000001 ↔ 1) ===
            prefix_match = False
            e_nums = re.findall(r"\d+", e_inv)
            v_nums = re.findall(r"\d+", v_inv)
            if e_nums and v_nums:
                e_last = e_nums[-1].lstrip("0") or "0"
                v_last = v_nums[-1].lstrip("0") or "0"
                prefix_match = e_last == v_last

            # === Weighted scoring ===
            score = 0
            if exact_match:
                score = 200
            elif last3_match:
                score = 150
            elif prefix_match:
                score = 130

            if score > best_score:
                best_score = score
                best_v = (v_idx, v_inv, v_amt, v_date)

        # === Confirm match if confidence >= 130 (Rule 3 or better) ===
        if best_v and best_score >= 130:
            v_idx, v_inv, v_amt, v_date = best_v
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
                "Status": status
            })

    # ===== Build DataFrames =====
    matched_df = pd.DataFrame(matched)

    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_vendor = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()

    # --- Missing in ERP ---
    missing_erp = ven_df[~ven_df["invoice_ven"].isin(matched_vendor)].copy()
    missing_erp = missing_erp.rename(columns={"invoice_ven": "Invoice"})
    missing_erp["Date"] = missing_erp.get("date_ven")
    missing_erp["Amount"] = missing_erp.get("debit_ven") or missing_erp.get("credit_ven")
    missing_erp = missing_erp[["Date", "Invoice", "Amount"]]

    # --- Missing in Vendor ---
    missing_vendor = erp_df[~erp_df["invoice_erp"].isin(matched_erp)].copy()
    missing_vendor = missing_vendor.rename(columns={"invoice_erp": "Invoice"})
    missing_vendor["Date"] = missing_vendor.get("date_erp")
    missing_vendor["Amount"] = missing_vendor.get("credit_erp") or missing_vendor.get("debit_erp")
    missing_vendor = missing_vendor[["Date", "Invoice", "Amount"]]

    return matched_df, missing_erp, missing_vendor



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
