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
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")

st.markdown("""
<style>
    .perfect-match { background-color: #2E7D32 !important; color: white !important; font-weight: bold !important; }
    .difference-match { background-color: #F9A825 !important; color: black !important; font-weight: bold !important; }
    .tier2-match { background-color: #26A69A !important; color: white !important; font-weight: bold !important; }
    .fuzzy-diff { background: linear-gradient(90deg, #FF5722, #FF8A65) !important; color: white !important; font-weight: bold !important; }
    .missing-erp { background-color: #C62828 !important; color: white !important; font-weight: bold !important; }
    .missing-vendor { background-color: #AD1457 !important; color: white !important; font-weight: bold !important; }
    .payment-match { background-color: #004D40 !important; color: white !important; font-weight: bold !important; }
    .erp-payment { background-color: #4CAF50 !important; color: white !important; }
    .vendor-payment { background-color: #2196F3 !important; color: white !important; }
    .metric-container { padding: 1rem !important; border-radius: 10px !important; }
    .total-row { background: linear-gradient(90deg, #667eea 0%, #764ba2 100%) !important; color: white !important; font-weight: bold !important; font-size: 14px !important; }
</style>
""", unsafe_allow_html=True)

st.title("🦖 ReconRaptor — Vendor Reconciliation")

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

def normalize_date(v):
    """Normalize date strings to YYYY-MM-DD format, handling various formats."""
    if pd.isna(v) or str(v).strip() == "":
        return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%m-%d-%Y",
        "%Y/%m/%d", "%Y-%m-%d",
        "%d/%m/%y", "%d-%m/%y", "%d.%m.%y",
        "%m/%d/%y", "%m-%d/%y",
        "%Y.%m.%d"
    ]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d):
                return d.strftime("%Y-%m-%d")
        except:
            continue
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(d):
            d = pd.to_datetime(s, errors="coerce", dayfirst=False)
        if pd.isna(d):
            return ""
        return d.strftime("%Y-%m-%d")
    except:
        return ""

def clean_invoice_code(v):
    """Clean invoice code to extract numerical components for fuzzy matching."""
    if not v:
        return ""
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
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document", "document number",
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου",
            "αριθμός τιμολογίου", "αριθμός παραστατικού", "κωδικός τιμολογίου", "τιμολόγιο", "αρ. παραστατικού",
            "παραστατικό τιμολογίου", "κωδικός παραστατικού"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού",
            "ποσό πίστωσης", "πιστωτικό ποσό"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            "χρέωση", "αξία", "αξία τιμολογίου",
            "ποσό χρέωσης", "συνολική αξία", "καθαρή αξία", "ποσό", "ποσό τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή",
            "description", "περιγραφή τιμολογίου", "αιτιολογία παραστατικού", "λεπτομέρειες"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού",
            "issue date", "transaction date", "emission date", "posting date",
            "ημερομηνία τιμολογίου", "ημερομηνία έκδοσης τιμολογίου", "ημερομηνία καταχώρισης",
            "ημερ. έκδοσης", "ημερ. παραστατικού", "ημερομηνία έκδοσης παραστατικού"
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
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

# ======================================
# SUMMARY TABLE WITH TOTALS
# ======================================
def create_summary_table_with_totals(matched_df, erp_missing, ven_missing):
    """Create summary table with all totals and differences"""
    erp_total = matched_df['ERP Amount'].sum() + erp_missing['Amount'].sum()
    vendor_total = matched_df['Vendor Amount'].sum() + ven_missing['Amount'].sum()
    matched_erp_total = matched_df['ERP Amount'].sum()
    matched_vendor_total = matched_df['Vendor Amount'].sum()
    total_difference = abs(erp_total - vendor_total)
   
    summary_data = {
        'Category': [
            '🔵 ERP Total Amount',
            '🟡 Vendor Total Amount',
            '🔴 Total Difference',
            '',
            '✅ Matched ERP Amount',
            '✅ Matched Vendor Amount',
            '⚠️ Matched Difference',
            '',
            '❌ Unmatched ERP',
            '❌ Unmatched Vendor'
        ],
        'Count': [
            len(matched_df) + len(erp_missing),
            len(matched_df) + len(ven_missing),
            '',
            '',
            len(matched_df),
            len(matched_df),
            '',
            '',
            len(erp_missing),
            len(ven_missing)
        ],
        'Amount': [
            f"€{erp_total:,.2f}",
            f"€{vendor_total:,.2f}",
            f"€{total_difference:,.2f}",
            '',
            f"€{matched_erp_total:,.2f}",
            f"€{matched_vendor_total:,.2f}",
            f"€{abs(matched_erp_total - matched_vendor_total):,.2f}",
            '',
            f"€{erp_missing['Amount'].sum():,.2f}",
            f"€{ven_missing['Amount'].sum():,.2f}"
        ]
    }
    return pd.DataFrame(summary_data)

def style_summary_table(df):
    """Style the summary table with colors"""
    def highlight_totals(row):
        if 'Total' in str(row['Category']):
            return ['background-color: linear-gradient(90deg, #667eea 0%, #764ba2 100%); color: white; font-weight: bold; font-size: 14px'] * len(row)
        elif 'Matched' in str(row['Category']):
            return ['background-color: #2E7D32; color: white; font-weight: bold'] * len(row)
        elif 'Unmatched' in str(row['Category']):
            return ['background-color: #C62828; color: white; font-weight: bold'] * len(row)
        else:
            return [''] * len(row)
    return df.style.apply(highlight_totals, axis=1)

# ======================================
# COLOR STYLING FUNCTIONS
# ======================================
def style_perfect_matches(df):
    return df.style.apply(lambda row: ['background-color: #2E7D32; color: white; font-weight: bold'] * len(row), axis=1)

def style_difference_matches(df):
    return df.style.apply(lambda row: ['background-color: #F9A825; color: black; font-weight: bold'] * len(row), axis=1)

def style_tier2_matches(df):
    return df.style.apply(lambda row: ['background-color: #26A69A; color: white; font-weight: bold'] * len(row), axis=1)

def style_fuzzy_diff_matches(df):
    return df.style.apply(lambda row: ['background: linear-gradient(90deg, #FF5722, #FF8A65); color: white; font-weight: bold'] * len(row), axis=1)

def style_missing(df):
    return df.style.apply(lambda row: ['background-color: #C62828; color: white; font-weight: bold'] * len(row), axis=1)

# ======================================
# CORE MATCHING - TIER 1
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()
  
    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        payment_patterns = [r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer",
                           r"^trf", r"^remesa", r"^pago", r"^pagado", r"^transferencia", r"^εξοφληση", r"^paid"]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in credit_words):
            return "CN"
        elif any(k in reason for k in invoice_words) or credit > 0:
            return "INV"
        return "UNKNOWN"
  
    def calc_erp_amount(row):
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if abs(charge) > 0:
            return abs(charge)
        elif abs(credit) > 0:
            return abs(credit)
        return 0.0
  
    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        payment_patterns = [r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer",
                           r"^trf", r"^remesa", r"^pago", r"^pagado", r"^transferencia", r"^εξοφληση", r"^paid"]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in credit_words) or credit > 0:
            return "CN"
        elif any(k in reason for k in invoice_words) or debit > 0:
            return "INV"
        return "UNKNOWN"
  
    def calc_vendor_amount(row):
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        if abs(debit) > 0:
            return abs(debit)
        elif abs(credit) > 0:
            return abs(credit)
        return 0.0
  
    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)
  
    erp_use = erp_df[erp_df["__doctype"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__doctype"] != "IGNORE"].copy()
  
    def merge_inv_cn(group_df, inv_col):
        merged_rows = []
        for inv, group in group_df.groupby(inv_col, dropna=False):
            if group.empty: continue
            if len(group) >= 3:
                group = group.tail(2)
            inv_rows = group[group["__doctype"] == "INV"]
            cn_rows = group[group["__doctype"] == "CN"]
            if not inv_rows.empty and not cn_rows.empty:
                total_inv = inv_rows["__amt"].sum()
                total_cn = cn_rows["__amt"].sum()
                net = round(abs(total_inv - total_cn), 2)
                base_row = inv_rows.iloc[-1].copy()
                base_row["__amt"] = net
                merged_rows.append(base_row)
            else:
                merged_rows.append(group.loc[group["__amt"].idxmax()])
        return pd.DataFrame(merged_rows).reset_index(drop=True)
  
    erp_use = merge_inv_cn(erp_use, "invoice_erp")
    ven_use = merge_inv_cn(ven_use, "invoice_ven")
  
    erp_use["__amt"] = erp_use["__amt"].astype(float)
    ven_use["__amt"] = ven_use["__amt"].astype(float)
  
    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_type = e["__doctype"]
      
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_type = v["__doctype"]
            diff = abs(e_amt - v_amt)
          
            if e_type != v_type: continue
            exact_match = (e_inv == v_inv)
            numerical_match = False
            e_nums = re.findall(r'(\d{4,})$', e_inv)
            v_nums = re.findall(r'(\d{4,})$', v_inv)
            if e_nums and v_nums and len(e_nums[0]) == len(v_nums[0]):
                numerical_match = (e_nums[0] == v_nums[0])
          
            amt_tolerance = 0.01
            amt_close = diff <= amt_tolerance
          
            if exact_match or numerical_match:
                if amt_close:
                    status = "Perfect Match"
                elif diff < 1.0:
                    status = "Difference Match"
                else:
                    continue
              
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": status
                })
                used_vendor_rows.add(v_idx)
                break
  
    matched_df = pd.DataFrame(matched)
    matched_erp = {m["ERP Invoice"] for _, m in matched_df.iterrows()}
    matched_ven = {m["Vendor Invoice"] for _, m in matched_df.iterrows()}
  
    erp_columns = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in erp_use.columns else [])
    ven_columns = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in ven_use.columns else [])
  
    missing_in_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)][erp_columns]
    missing_in_vendor = ven_use[~ven_use["invoice_ven"].isin(matched_erp)][ven_columns]
  
    missing_in_erp = missing_in_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
  
    return matched_df, missing_in_erp, missing_in_vendor

# ======================================
# TIER-2 MATCHING (FUZZY + EXACT AMOUNT)
# ======================================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def tier2_match(erp_missing, ven_missing):
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), set(), set(), erp_missing.copy(), ven_missing.copy()
    
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    
    e_df["date_norm"] = e_df["date_erp"].apply(normalize_date) if "date_erp" in e_df.columns else ""
    v_df["date_norm"] = v_df["date_ven"].apply(normalize_date) if "date_ven" in v_df.columns else ""
    
    matches, used_e, used_v = [], set(), set()
    
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e: continue
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0)), 2)
        e_date = e.get("date_norm", "")
        e_code = clean_invoice_code(e_inv)
        
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0)), 2)
            v_date = v.get("date_norm", "")
            v_code = clean_invoice_code(v_inv)
            
            diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_code, v_code)
            
            if diff < 0.05 and sim >= 0.8:
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Fuzzy Score": round(sim, 2),
                    "Date": e_date or v_date or "N/A",
                    "Match Type": "Tier-2"
                })
                used_e.add(e_idx)
                used_v.add(v_idx)
                break
    
    tier2_matches = pd.DataFrame(matches)
    erp_columns = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in e_df.columns else [])
    ven_columns = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in v_df.columns else [])
    
    remaining_erp_missing = e_df[~e_df.index.isin(used_e)][erp_columns].rename(
        columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"}
    )
    remaining_ven_missing = v_df[~v_df.index.isin(used_v)][ven_columns].rename(
        columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"}
    )
    
    return tier2_matches, used_e, used_v, remaining_erp_missing, remaining_ven_missing

# ======================================
# TIER-3: FUZZY MATCHES WITH DIFFERENT AMOUNTS (NEW!)
# ======================================
def fuzzy_different_amounts_match(erp_missing, ven_missing):
    """Fuzzy matching for invoices with DIFFERENT amounts (5-20% tolerance)"""
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), set(), set(), erp_missing.copy(), ven_missing.copy()
    
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    
    e_df["date_norm"] = e_df["date_erp"].apply(normalize_date) if "date_erp" in e_df.columns else ""
    v_df["date_norm"] = v_df["date_ven"].apply(normalize_date) if "date_ven" in v_df.columns else ""
    
    matches, used_e, used_v = [], set(), set()
    
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e: continue
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0)), 2)
        e_date = e.get("date_norm", "")
        e_code = clean_invoice_code(e_inv)
        
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0)), 2)
            v_date = v.get("date_norm", "")
            v_code = clean_invoice_code(v_inv)
            
            # Calculate percentage difference
            pct_diff = abs(e_amt - v_amt) / max(e_amt, v_amt) * 100 if max(e_amt, v_amt) > 0 else 0
            abs_diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_code, v_code)
            
            # FUZZY MATCH CRITERIA: High similarity + 5-20% amount difference
            if sim >= 0.85 and 5 <= pct_diff <= 20:
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Absolute Diff": abs_diff,
                    "Pct Diff (%)": round(pct_diff, 1),
                    "Fuzzy Score": round(sim, 3),
                    "Date": e_date or v_date or "N/A",
                    "Match Type": "Fuzzy Diff"
                })
                used_e.add(e_idx)
                used_v.add(v_idx)
                break
    
    fuzzy_diff_matches = pd.DataFrame(matches)
    erp_columns = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in e_df.columns else [])
    ven_columns = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in v_df.columns else [])
    
    remaining_erp_missing = e_df[~e_df.index.isin(used_e)][erp_columns].rename(
        columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"}
    )
    remaining_ven_missing = v_df[~v_df.index.isin(used_v)][ven_columns].rename(
        columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"}
    )
    
    return fuzzy_diff_matches, used_e, used_v, remaining_erp_missing, remaining_ven_missing

# ======================================
# PAYMENTS
# ======================================
def extract_payments(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    payment_keywords = ["πληρωμή", "payment", "bank transfer", "transferencia bancaria",
                       "transfer", "trf", "remesa", "pago", "deposit", "μεταφορά", "έμβασμα",
                       "εξοφληση", "pagado", "paid"]
    exclude_keywords = ["invoice of expenses", "expense invoice", "τιμολόγιο εξόδων",
                       "διόρθωση", "διορθώσεις", "correction", "reclass", "adjustment",
                       "μεταφορά υπολοίπου", "balance transfer"]
    
    def is_real_payment(row: pd.Series, tag: str) -> bool:
        text = str(row.get(f"reason_{tag}", "")).lower()
        has_payment = any(k in text for k in payment_keywords)
        has_exclusion = any(bad in text for bad in exclude_keywords)
        if tag == "erp":
            debit = normalize_number(row.get("debit_erp", 0))
            return has_payment and not has_exclusion and debit > 0
        elif tag == "ven":
            credit = normalize_number(row.get("credit_ven", 0))
            return has_payment and not has_exclusion and credit > 0
        return False
    
    erp_pay = erp_df[erp_df.apply(lambda x: is_real_payment(x, "erp"), axis=1)].copy() if "reason_erp" in erp_df else pd.DataFrame()
    ven_pay = ven_df[ven_df.apply(lambda x: is_real_payment(x, "ven"), axis=1)].copy() if "reason_ven" in ven_df else pd.DataFrame()
    
    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(lambda r: abs(normalize_number(r.get("debit_erp
