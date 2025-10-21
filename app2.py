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
            "invoice","factura","fact","nº","num","numero","número","document","doc","ref",
            "referencia","nº factura","num factura","alternative document","αρ.","αριθμός",
            "νουμερο","νούμερο","no","παραστατικό","αρ. τιμολογίου","αρ. εγγράφου"
        ],
        "credit": [
            "credit","haber","credito","crédito","nota de crédito","nota crédito","abono","abonos",
            "importe haber","valor haber","πίστωση","πιστωτικό","πιστωτικό τιμολόγιο","πίστωση ποσού"
        ],
        "debit": [
            "debit","debe","cargo","importe","importe total","valor","μonto","amount",
            "document value","charge","total","totale","totales","totals","base imponible",
            "importe factura","importe neto","χρέωση","αξία","αξία τιμολογίου"
        ],
        "reason": [
            "reason","motivo","concepto","descripcion","descripción","detalle","detalles","razon",
            "razón","observaciones","comentario","comentarios","explicacion","αιτιολογία","περιγραφή",
            "παρατηρήσεις","σχόλια","αναφορά","αναλυτική περιγραφή"
        ],
        "cif": [
            "cif","nif","vat","iva","tax","id fiscal","número fiscal","num fiscal","code","αφμ",
            "φορολογικός αριθμός","αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date","fecha","fech","data","fecha factura","fecha doc","fecha documento","ημερομηνία",
            "ημ/νία","ημερομηνία έκδοσης","ημερομηνία παραστατικού"
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
            r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer",
            r"^trf", r"^remesa", r"^pago", r"^transferencia", r"έμβασμα\s*από\s*πελάτη\s*χειρ."
        ]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit","nota","abono","cn","πιστωτικό","πίστωση","ακυρωτικό"]
        invoice_words = ["factura","invoice","inv","τιμολόγιο","παραστατικό"]
        if any(k in reason for k in credit_words):
            return "CN"
        elif any(k in reason for k in invoice_words) or credit > 0:
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
        payment_words = ["pago","payment","transfer","bank","saldo","trf","πληρωμή","μεταφορά","τράπεζα","τραπεζικό έμβασμα"]
        credit_words = ["credit","nota","abono","cn","πιστωτικό","πίστωση","ακυρωτικό"]
        invoice_words = ["factura","invoice","inv","τιμολόγιο","παραστατικό"]
        if any(k in reason for k in payment_words):
            return "IGNORE"
        elif any(k in reason for k in credit_words) or credit > 0:
            return "CN"
        elif any(k in reason for k in invoice_words) or debit > 0:
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
        if not v:
            return ""
        s = str(v).strip().lower()
        s = re.sub(r"[^a-z0-9]", "", s)
        return s

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_code = clean_invoice_code(e_inv)
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_code = clean_invoice_code(v_inv)
            diff = round(e_amt - v_amt, 2)
            amt_close = abs(diff) < 0.05
            same_type = (e["__doctype"] == v["__doctype"])
            same_clean = (e_code == v_code)
            if same_type and same_clean:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match" if amt_close else "Difference"
                })
                used_vendor_rows.add(v_idx)
                break

    # ✅ Safe handling
    if matched:
        matched_df = pd.DataFrame(matched)
        matched_erp = set(matched_df["ERP Invoice"])
        matched_ven = set(matched_df["Vendor Invoice"])
    else:
        matched_df = pd.DataFrame(columns=["ERP Invoice","Vendor Invoice","ERP Amount","Vendor Amount","Difference","Status"])
        matched_erp, matched_ven = set(), set()

    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]]
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]]

    missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    return matched_df, missing_in_erp, missing_in_vendor
