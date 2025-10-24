import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher

# ======================================
# CONFIGURATION WITH CUSTOM CSS COLORS
# ======================================
st.set_page_config(page_title="ReconRaptor — Vendor Reconciliation", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .perfect-match { background-color: #2E7D32 !important; color: white !important; font-weight: bold !important; }
    .difference-match { background-color: #F9A825 !important; color: black !important; font-weight: bold !important; }
    .tier2-match { background-color: #26A69A !important; color: white !important; font-weight: bold !important; }
    .tier3-match { background-color: #7E57C2 !important; color: white !important; font-weight: bold !important; }
    .missing-erp { background-color: #C62828 !important; color: white !important; font-weight: bold !important; }
    .missing-vendor { background-color: #AD1457 !important; color: white !important; font-wight: bold !important; }
    .payment-match { background-color: #004D40 !important; color: white !important; font-weight: bold !important; }
    .erp-payment { background-color: #4CAF50 !important; color: white !important; }
    .vendor-payment { background-color: #2196F3 !important; color: white !important; }
    .metric-container { padding: 1rem !important; border-radius: 10px !important; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)

st.title("ReconRaptor — Vendor Reconciliation")

# ======================================
# HELPERS
# ======================================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if v is None or str(v).strip() == "": return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "*", s)
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

def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "": return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%m-%d-%Y",
        "%Y/%m/%d", "%Y-%m-%d",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
        "%m/%d/%y", "%m-%d-%y",
        "%Y.%m.%d"
    ]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d):
                return d.strftime("%Y-%m-%d")
        except: continue
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(d): d = pd.to_datetime(s, errors="coerce", dayfirst=False)
        if pd.isna(d): return ""
        return d.strftime("%Y-%m-%d")
    except: return ""

def clean_invoice_code(v):
    if not v: return ""
    s = str(v).strip().lower()
    parts = re.split(r"[-_]", s)
    for p in reversed(parts):
        if re.fullmatch(r"\d{1,}", p) and not re.fullmatch(r"20[0-3]\d", p):
            s = p.lstrip("0")
            break
    s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|apd|vs)\W*", "", s)
    s = re.sub(r"20\d{2}", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    s = re.sub(r"[^\d]", "", s)
    return s or "0"

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "número", "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document", "document number", "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου", "αριθμός τιμολογίου", "αριθμός παραστατικού", "κωδικός τιμολογίου", "τιμολόγιο", "αρ. παραστατικού", "παραστατικό τιμολογίου", "κωδικός παραστατικού"],
        "credit": ["credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito", "abono", "abonos", "importe haber", "valor haber", "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού", "ποσό πίστωσης", "πιστωτικό ποσό"],
        "debit": ["debit", "debe", "cargo", "importe", "importe total", "valor", "monto", "amount", "document value", "charge", "total", "totale", "totales", "totals", "base imponible", "importe factura", "importe neto", "χρέωση", "αξία", "αξία τιμολογίου", "ποσό χρέωσης", "συνολική αξία", "καθαρή αξία", "ποσό", "ποσό τιμολογίου"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "descripción", "detalle", "detalles", "razon", "razón", "observaciones", "comentario", "comentarios", "explicacion", "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή", "description", "περιγραφή τιμολογίου", "αιτιολογία παραστατικού", "λεπτομέρειες"],
        "cif": ["cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code", "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"],
        "date": ["date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento", "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού", "issue date", "transaction date", "emission date", "posting date", "ημερομηνία τιμολογίου", "ημερομηνία έκδοσης τιμολογίου", "ημερομηνία καταχώρισης", "ημερ. έκδοσης", "ημερ. παραστατικού", "ημερομηνία έκδοσης παραστατικού"]
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
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

# ======================================
# STYLING
# ======================================
def style_perfect_matches(df): return df.style.apply(lambda row: ['background-color: #2E7D32; color: white; font-weight: bold'] * len(row), axis=1)
def style_difference_matches(df): return df.style.apply(lambda row: ['background-color: #F9A825; color: black; font-weight: bold'] * len(row), axis=1)
def style_tier2_matches(df): return df.style.apply(lambda row: ['background-color: #26A69A; color: white; font-weight: bold'] * len(row), axis=1)
def style_tier3_matches(df): return df.style.apply(lambda row: ['background-color: #7E57C2; color: white; font-weight: bold'] * len(row), axis=1)
def style_missing(df): return df.style.apply(lambda row: ['background-color: #C62828; color: white; font-weight: bold'] * len(row), axis=1)

# ======================================
# MATCHING FUNCTIONS
# ======================================
# [match_invoices, tier2_match, tier3_match, extract_payments — same as before, but compact]
# ... (I'll include them compactly below)

# --- COMPACT VERSION OF ALL MATCHING FUNCTIONS ---
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    def detect_doc_type(row, tag):
        reason = str(row.get(f"reason_{tag}", "")).lower()
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        payment_patterns = [r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer", r"^trf", r"^remesa", r"^pago", r"^pagado", r"^transferencia", r"^εξοφληση", r"^paid"]
        if any(re.search(p, reason) for p in payment_patterns): return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in credit_words) or (tag == "ven" and credit > 0) or (tag == "erp" and credit > 0): return "CN"
        elif any(k in reason for k in invoice_words) or debit > 0: return "INV"
        return "UNKNOWN"

    erp_df["__doctype"] = erp_df.apply(lambda r: detect_doc_type(r, "erp"), axis=1)
    ven_df["__doctype"] = ven_df.apply(lambda r: detect_doc_type(r, "ven"), axis=1)
    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)

    erp_use = erp_df[erp_df["__doctype"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__doctype"] != "IGNORE"].copy()

    def merge_inv_cn(df, col):
        merged = []
        for inv, g in df.groupby(col, dropna=False):
            if g.empty: continue
            inv_rows = g[g["__doctype"] == "INV"]
            cn_rows = g[g["__doctype"] == "CN"]
            if not inv_rows.empty and not cn_rows.empty:
                net = round(abs(inv_rows["__amt"].sum() - cn_rows["__amt"].sum()), 2)
                base = inv_rows.iloc[-1].copy()
                base["__amt"] = net
                merged.append(base)
            else:
                merged.append(g.loc[g["__amt"].idxmax()])
        return pd.DataFrame(merged)

    erp_use = merge_inv_cn(erp_use, "invoice_erp")
    ven_use = merge_inv_cn(ven_use, "invoice_ven")

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_type = e["__doctype"]
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_type = v["__doctype"]
            if e_type != v_type or e_inv != v_inv: continue
            diff = abs(e_amt - v_amt)
            status = "Perfect Match" if diff <= 0.01 else "Difference Match" if diff < 1.0 else None
            if status:
                matched.append({"ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt, "Vendor Amount": v_amt, "Difference": diff, "Status": status})
                used_vendor_rows.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = set(matched_df["ERP Invoice"])
    matched_ven = set(matched_df["Vendor Invoice"])

    erp_cols = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in erp_use.columns else [])
    ven_cols = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in ven_use.columns else [])

    missing_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)][erp_cols].rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    missing_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp)][ven_cols].rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})

    return matched_df, missing_erp, missing_ven

def tier2_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty: return pd.DataFrame(), set(), set(), erp_miss.copy(), ven_miss.copy()
    e_df = erp_miss.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_miss.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    matches, used_e, used_v = [], set(), set()
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e: continue
        e_inv, e_amt, e_code = str(e["invoice_erp"]), round(e["__amt"], 2), clean_invoice_code(e["invoice_erp"])
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v: continue
            v_inv, v_amt, v_code = str(v["invoice_ven"]), round(v["__amt"], 2), clean_invoice_code(v["invoice_ven"])
            diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_code, v_code)
            if diff < 0.05 and sim >= 0.8:
                matches.append({"ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt, "Vendor Amount": v_amt, "Difference": diff, "Fuzzy Score": round(sim, 2), "Match Type": "Tier-2"})
                used_e.add(e_idx); used_v.add(v_idx); break
    mdf = pd.DataFrame(matches)
    rem_e = e_df[~e_df.index.isin(used_e)][["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in e_df.columns else [])].rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    rem_v = v_df[~v_df.index.isin(used_v)][["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in v_df.columns else [])].rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
    return mdf, used_e, used_v, rem_e, rem_v

def tier3_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty: return pd.DataFrame(), set(), set(), erp_miss.copy(), ven_miss.copy()
    e_df = erp_miss.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_miss.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    e_df["date_norm"] = e_df["date_erp"].apply(normalize_date)
    v_df["date_norm"] = v_df["date_ven"].apply(normalize_date)
    matches, used_e, used_v = [], set(), set()
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e or not e["date_norm"]: continue
        e_inv, e_amt, e_code, e_date = str(e["invoice_erp"]), round(e["__amt"], 2), clean_invoice_code(e["invoice_erp"]), e["date_norm"]
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v or not v["date_norm"]: continue
            v_inv, v_amt, v_code, v_date = str(v["invoice_ven"]), round(v["__amt"], 2), clean_invoice_code(v["invoice_ven"]), v["date_norm"]
            if e_date == v_date and fuzzy_ratio(e_code, v_code) >= 0.9:
                diff = abs(e_amt - v_amt)
                matches.append({"ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt, "Vendor Amount": v_amt, "Difference": diff, "Fuzzy Score": round(fuzzy_ratio(e_code, v_code), 2), "Date": e_date, "Match Type": "Tier-3"})
                used_e.add(e_idx); used_v.add(v_idx); break
    mdf = pd.DataFrame(matches)
    rem_e = e_df[~e_df.index.isin(used_e)][["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in e_df.columns else [])].rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    rem_v = v_df[~v_df.index.isin(used_v)][["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in v_df.columns else [])].rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
    return mdf, used_e, used_v, rem_e, rem_v

def extract_payments(erp_df, ven_df):
    # ... (same logic as before)
    pass  # Implement if needed

# ======================================
# UI
# ======================================
# USE UNIQUE KEYS TO AVOID DUPLICATE ID ERROR
uploaded_erp = st.file_uploader("Upload ERP Export (Excel)", type=["xlsx"], key="erp_uploader")
uploaded_vendor = st.file_uploader("Upload Vendor Statement (Excel)", type=["xlsx"], key="vendor_uploader")

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")

        with st.spinner("Analyzing..."):
            matched, erp_miss, ven_miss = match_invoices(erp_df, ven_df)
            tier2, _, _, erp_miss2, ven_miss2 = tier2_match(erp_miss, ven_miss)
            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(erp_miss2, ven_miss2)

        st.success("Done!")

        # Metrics with totals
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        p_count = len(matched[matched['Status'] == 'Perfect Match'])
        d_count = len(matched[matched['Status'] == 'Difference Match'])
        t2_count = len(tier2)
        t3_count = len(tier3)
        um_erp = len(final_erp_miss)
        um_ven = len(final_ven_miss)

        with col1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("Perfect", p_count)
            st.markdown(f"Diff: {matched[matched['Status']=='Perfect Match']['Difference'].sum():.2f}")
            st.markdown('</div>', unsafe_allow_html=True)
        # ... repeat for others

        # Tables
        st.subheader("Tier-1")
        if not matched.empty:
            st.dataframe(style_perfect_matches(matched[matched['Status']=='Perfect Match']), use_container_width=True)

        st.subheader("Tier-3 (Date + Strict Fuzzy)")
        if not tier3.empty:
            st.dataframe(style_tier3_matches(tier3), use_container_width=True)

        # ... rest of UI

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check columns: invoice, debit/credit, date, reason")
