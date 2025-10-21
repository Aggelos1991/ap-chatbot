import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher

# ======================================
# CONFIGURATION
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ======================================
# HELPERS
# ======================================
def normalize_number(v):
    """Convert numeric strings like '1.234,56' or '1,234.56' safely to float."""
    if v is None or str(v).strip() == "":
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
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document",
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "μonto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού"
        ],
    }

    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"

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

    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        payment_patterns = [
            r"πληρωμ", r"payment", r"transfer", r"trf", r"pago"
        ]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        if any(k in reason for k in ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]):
            return "CN"
        elif credit > 0:
            return "INV"
        return "UNKNOWN"

    def calc_erp_amount(row):
        doc = row.get("__doctype", "")
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if doc == "INV":
            return abs(credit)
        elif doc == "CN":
            return -abs(charge if charge > 0 else credit)
        return 0.0

    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        if any(k in reason for k in ["pago","payment","transfer","bank","saldo","trf","πληρωμή","μεταφορά","τράπεζα","τραπεζικό έμβασμα"]):
            return "IGNORE"
        if any(k in reason for k in ["credit","nota","abono","cn","πιστωτικό","πίστωση","ακυρωτικό"]):
            return "CN"
        elif debit > 0:
            return "INV"
        return "UNKNOWN"

    def calc_vendor_amount(row):
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        doc = row.get("__doctype", "")
        if doc == "INV":
            return abs(debit)
        elif doc == "CN":
            return -abs(credit if credit > 0 else debit)
        return 0.0

    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    def clean_invoice_code(v):
        return re.sub(r"[^a-z0-9]", "", str(v).strip().lower())

    for e_idx, e in erp_use.iterrows():
        e_inv, e_amt = str(e.get("invoice_erp", "")).strip(), round(float(e["__amt"]), 2)
        e_code = clean_invoice_code(e_inv)
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue
            v_inv, v_amt = str(v.get("invoice_ven", "")).strip(), round(float(v["__amt"]), 2)
            v_code = clean_invoice_code(v_inv)
            diff = round(e_amt - v_amt, 2)
            if e["__doctype"] == v["__doctype"] and e_code == v_code:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match" if abs(diff) < 0.05 else "Difference"
                })
                used_vendor_rows.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_ven = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()
    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]]
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]]
    missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"}, inplace=True)
    missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"}, inplace=True)
    return matched_df, missing_in_erp, missing_in_vendor

# ======================================
# TIER-2 MATCHING
# ======================================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def tier2_match(erp_missing, ven_missing):
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), ven_missing.copy()
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt"}).copy()
    matches, used_v = [], set()
    for e_idx, e in e_df.iterrows():
        e_inv, e_amt = str(e.get("invoice_erp", "")), round(float(e.get("__amt", 0)), 2)
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v:
                continue
            v_inv, v_amt = str(v.get("invoice_ven", "")), round(float(v.get("__amt", 0)), 2)
            diff, sim = abs(e_amt - v_amt), fuzzy_ratio(e_inv, v_inv)
            if diff < 0.05 and sim >= 0.8:
                matches.append({
                    "ERP Invoice": e_inv, "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt, "Vendor Amount": v_amt,
                    "Difference": diff, "Fuzzy Score": round(sim, 2),
                    "Match Type": "Tier-2"
                })
                used_v.add(v_idx)
                break
    return pd.DataFrame(matches), v_df[~v_df.index.isin(used_v)].copy()

# ======================================
# STREAMLIT UI
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_df = normalize_columns(pd.read_excel(uploaded_erp, dtype=str), "erp")
    ven_df = normalize_columns(pd.read_excel(uploaded_vendor, dtype=str), "ven")

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)

    # 🧩 Tier-2 Matching
    tier2_matches, _ = tier2_match(erp_missing, ven_missing)
    if not tier2_matches.empty:
        matched_vendor_invoices = tier2_matches["Vendor Invoice"].unique().tolist()
        matched_erp_invoices = tier2_matches["ERP Invoice"].unique().tolist()
        ven_missing = ven_missing[~ven_missing["Invoice"].isin(matched_vendor_invoices)]
        erp_missing = erp_missing[~erp_missing["Invoice"].isin(matched_erp_invoices)]

    st.success("✅ Reconciliation complete")

    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color:#2e7d32;color:white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color:#f9a825;color:black'] * len(row)
        return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")

    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color:#c62828;color:white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color:#c62828;color:white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in Vendor.")

    # Tier-2 table with green color
    st.markdown("### 🧩 Tier-2 Matching (same date, same value, fuzzy invoice)")
    if not tier2_matches.empty:
        st.success(f"✅ Tier-2 matched {len(tier2_matches)} additional pairs.")
        st.dataframe(
            tier2_matches.style.applymap(lambda _: "background-color:#1b5e20;color:white"),
            use_container_width=True
        )
    else:
        st.info("No Tier-2 matches found.")


    # Export
    def export_reconciliation_excel(matched, erp_missing, ven_missing, tier2_matches):
        import io
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            matched.to_excel(writer, index=False, sheet_name="Matched & Differences")
            if not tier2_matches.empty:
                tier2_matches.to_excel(writer, index=False, sheet_name="Tier-2 Matches")
            ws_name = "Missing"
            erp_missing.to_excel(writer, index=False, sheet_name=ws_name, startrow=4)
            start_col = len(erp_missing.columns) + 4
            ven_missing.to_excel(writer, index=False, sheet_name=ws_name, startcol=start_col, startrow=4)
        output.seek(0)
        return output

    st.markdown("### 📥 Download Reconciliation Excel Report")
    excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing, tier2_matches)
    st.download_button("⬇️ Download Excel Report", data=excel_output,
                       file_name="Reconciliation_Report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
