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
# Custom CSS for beautiful color styling
st.markdown("""
<style>
    .perfect-match { background-color: #2E7D32 !important; color: white !important; font-weight: bold !important; }
    .difference-match { background-color: #F9A825 !important; color: black !important; font-weight: bold !important; }
    .tier2-match { background-color: #26A69A !important; color: white !important; font-weight: bold !important; }
    .tier2-amount-diff { background-color: #FFCA28 !important; color: black !important; font-weight: bold !important; }
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
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
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
def clean_invoice_code_strict(v):
    """Clean invoice code strictly by removing special characters, years, slashes, spaces, and non-numeric parts."""
    if not v:
        return ""
    s = str(v).strip().lower()
    s = re.sub(r"[^a-z0-9]", "", s) # Keep only alphanumeric
    s = re.sub(r"20\d{2}", "", s) # Remove year patterns like 2023
    s = re.sub(r"^\D+", "", s) # Remove leading non-numeric characters
    s = re.sub(r"^0+", "", s) # Remove leading zeros
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
        "due_date": [
            "due date", "fecha vencimiento", "vencimiento", "due", "payable by",
            "ημερομηνία λήξης", "λήξη", "ημ/νία λήξης", "due date",
            "payment due", "expiration date"
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
    for date_col in [f"date_{tag}", f"due_date_{tag}"]:
        if date_col in out.columns:
            out[date_col] = out[date_col].apply(normalize_date)
    return out
# ======================================
# MISSING FUNCTIONS IMPLEMENTED
# ======================================
def extract_payments(erp_df, ven_df):
    """Extract and match payment records from both datasets."""
    def detect_payments(df, tag):
        payments = []
        payment_patterns = [
            r"πληρωμ", r"απόδειξη\s*πληρωμ", r"payment", r"bank\s*transfer",
            r"trf", r"remesa", r"pago", r"pagado", r"transferencia",
            r"εξοφληση", r"paid", r"settled", r"clearing"
        ]
    
        for idx, row in df.iterrows():
            reason = str(row.get(f"reason_{tag}", "")).lower()
            if any(re.search(p, reason) for p in payment_patterns):
                payments.append({
                    "index": idx,
                    "reason": row.get(f"reason_{tag}", ""),
                    "amount": abs(normalize_number(row.get(f"debit_{tag}", row.get(f"credit_{tag}", 0)))),
                    "date": row.get(f"date_{tag}", "")
                })
        return pd.DataFrame(payments)
    erp_payments = detect_payments(erp_df, "erp")
    ven_payments = detect_payments(ven_df, "ven")
    matched_payments = pd.DataFrame()
    return erp_payments, ven_payments, matched_payments
def export_reconciliation_excel(matched, erp_missing, ven_missing, matched_payments, tier2_matches, tier2_amount_diff):
    """Export complete reconciliation report to Excel."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not matched.empty:
            matched.to_excel(writer, sheet_name='Matched_Invoices', index=False)
        if not erp_missing.empty:
            erp_missing.to_excel(writer, sheet_name='ERP_Missing', index=False)
        if not ven_missing.empty:
            ven_missing.to_excel(writer, sheet_name='Vendor_Missing', index=False)
        if not tier2_matches.empty:
            tier2_matches.to_excel(writer, sheet_name='Tier2_Matches', index=False)
        if not tier2_amount_diff.empty:
            tier2_amount_diff.to_excel(writer, sheet_name='Tier2_Amount_Diff', index=False)
        if not matched_payments.empty:
            matched_payments.to_excel(writer, sheet_name='Matched_Payments', index=False)
        summary_data = {
            'Category': ['Matched Count', 'Tier-2 Matches', 'Tier-2 Amount Diff', 'ERP Unmatched', 'Vendor Unmatched'],
            'Count': [len(matched), len(tier2_matches), len(tier2_amount_diff), len(erp_missing), len(ven_missing)]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    output.seek(0)
    return output.getvalue()
# ======================================
# SUMMARY TABLE WITH TOTALS
# ======================================
def create_summary_table_with_totals(matched_df, tier2_matches, tier2_amount_diff, erp_missing, ven_missing):
    """Create summary table with all totals and differences."""
    erp_total = matched_df['ERP Amount'].sum() + tier2_matches['ERP Amount'].sum() + tier2_amount_diff['ERP Amount'].sum() + erp_missing['Amount'].sum()
    vendor_total = matched_df['Vendor Amount'].sum() + tier2_matches['Vendor Amount'].sum() + tier2_amount_diff['Vendor Amount'].sum() + ven_missing['Amount'].sum()
    total_difference = abs(erp_total - vendor_total)
    tier1_erp_total = matched_df['ERP Amount'].sum()
    tier1_vendor_total = matched_df['Vendor Amount'].sum()
    tier2_erp_total = tier2_matches['ERP Amount'].sum()
    tier2_vendor_total = tier2_matches['Vendor Amount'].sum()
    diff_erp_total = tier2_amount_diff['ERP Amount'].sum()
    diff_vendor_total = tier2_amount_diff['Vendor Amount'].sum()
    summary_data = {
        'Category': [
            'ERP Total Amount', 'Vendor Total Amount', 'Total Difference', '',
            'Tier-1 Matched ERP', 'Tier-1 Matched Vendor', 'Tier-1 Matched Difference', '',
            'Tier-2 Matched ERP', 'Tier-2 Matched Vendor', 'Tier-2 Matched Difference', '',
            'Tier-2 Amount Diff ERP', 'Tier-2 Amount Diff Vendor', 'Tier-2 Amount Diff Difference', '',
            'Unmatched ERP', 'Unmatched Vendor'
        ],
        'Count': [
            len(matched_df) + len(tier2_matches) + len(tier2_amount_diff) + len(erp_missing),
            len(matched_df) + len(tier2_matches) + len(tier2_amount_diff) + len(ven_missing),
            '', '', len(matched_df), len(matched_df), '', '',
            len(tier2_matches), len(tier2_matches), '', '',
            len(tier2_amount_diff), len(tier2_amount_diff), '', '',
            len(erp_missing), len(ven_missing)
        ],
        'Amount': [
            f"{erp_total:,.2f}", f"{vendor_total:,.2f}", f"{total_difference:,.2f}", '',
            f"{tier1_erp_total:,.2f}", f"{tier1_vendor_total:,.2f}", f"{abs(tier1_erp_total - tier1_vendor_total):,.2f}", '',
            f"{tier2_erp_total:,.2f}", f"{tier2_vendor_total:,.2f}", f"{abs(tier2_erp_total - tier2_vendor_total):,.2f}", '',
            f"{diff_erp_total:,.2f}", f"{diff_vendor_total:,.2f}", f"{abs(diff_erp_total - diff_vendor_total):,.2f}", '',
            f"{erp_missing['Amount'].sum():,.2f}", f"{ven_missing['Amount'].sum():,.2f}"
        ]
    }
    return pd.DataFrame(summary_data)
def style_summary_table(df):
    """Style the summary table with colors."""
    def highlight_totals(row):
        category = str(row['Category'])
        if 'Total' in category:
            return ['background-color: linear-gradient(90deg, #667eea 0%, #764ba2 100%); color: white; font-weight: bold; font-size: 14px'] * len(row)
        elif 'Tier-1 Matched' in category:
            return ['background-color: #2E7D32; color: white; font-weight: bold'] * len(row)
        elif 'Tier-2 Matched' in category:
            return ['background-color: #26A69A; color: white; font-weight: bold'] * len(row)
        elif 'Tier-2 Amount Diff' in category:
            return ['background-color: #FFCA28; color: black; font-weight: bold'] * len(row)
        elif 'Unmatched' in category:
            return ['background-color: #C62828; color: white; font-weight: bold'] * len(row)
        else:
            return [''] * len(row)
    return df.style.apply(highlight_totals, axis=1)
# ======================================
# CORE FUNCTIONS
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
            r"^trf", r"^remesa", r"^pago", r"^pagado", r"^transferencia",
            r"^εξοφληση", r"^paid",
        ]
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
        payment_patterns = [
            r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer",
            r"^trf", r"^remesa", r"^pago", r"^pagado", r"^transferencia",
            r"^εξοφληση", r"^paid",
        ]
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
            if v_idx in used_vendor_rows:
                continue
           
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_type = v["__doctype"]
       
            diff = abs(e_amt - v_amt)
       
            if e_type != v_type:
                continue
           
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
    erp_columns = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in erp_use.columns else []) + (["due_date_erp"] if "due_date_erp" in erp_use.columns else [])
    ven_columns = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in ven_use.columns else []) + (["due_date_ven"] if "due_date_ven" in ven_use.columns else [])
    missing_in_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)][erp_columns]
    missing_in_vendor = ven_use[~ven_use["invoice_ven"].isin(matched_erp)][ven_columns]
    missing_in_erp = missing_in_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date", "due_date_erp": "Due Date"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date", "due_date_ven": "Due Date"})
    return matched_df, missing_in_erp, missing_in_vendor
def strict_match_unmatched(erp_missing, ven_missing):
    """Match unmatched invoices strictly by cleaned invoice numbers, ignoring amounts."""
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), erp_missing.copy(), ven_missing.copy()
   
    # Clean invoice numbers strictly
    erp_missing['Cleaned_Invoice'] = erp_missing['Invoice'].apply(clean_invoice_code_strict)
    ven_missing['Cleaned_Invoice'] = ven_missing['Invoice'].apply(clean_invoice_code_strict)
   
    # Create dictionaries mapping cleaned invoices to their indices
    erp_dict = dict(zip(erp_missing['Cleaned_Invoice'], erp_missing.index))
    ven_dict = dict(zip(ven_missing['Cleaned_Invoice'], ven_missing.index))
   
    matches = []
    used_erp_indices = set()
    used_ven_indices = set()
   
    for idx_erp, cleaned_erp in enumerate(erp_missing['Cleaned_Invoice']):
        if idx_erp in used_erp_indices:
            continue
        if cleaned_erp in ven_dict:
            idx_ven = ven_dict[cleaned_erp]
            if 0 <= idx_ven < len(ven_missing) and idx_ven not in used_ven_indices:
                erp_row = erp_missing.iloc[idx_erp]
                ven_row = ven_missing.iloc[idx_ven]
                diff = abs(erp_row['Amount'] - ven_row['Amount'])
                matches.append({
                    "ERP Invoice": erp_row['Invoice'],
                    "Vendor Invoice": ven_row['Invoice'],
                    "ERP Amount": erp_row['Amount'],
                    "Vendor Amount": ven_row['Amount'],
                    "Difference": diff,
                    "Match Type": "Strict Match"
                })
                used_erp_indices.add(idx_erp)
                used_ven_indices.add(idx_ven)
   
    strict_matches = pd.DataFrame(matches)
   
    # Remove matched rows from missing datasets
    remaining_erp_missing = erp_missing.drop(used_erp_indices).drop(columns=['Cleaned_Invoice'])
    remaining_ven_missing = ven_missing.drop(used_ven_indices).drop(columns=['Cleaned_Invoice'])
   
    # Ensure Date and Due Date columns exist, even if empty
    for col in ['Date', 'Due Date']:
        if col not in remaining_erp_missing.columns:
            remaining_erp_missing[col] = ""
        if col not in remaining_ven_missing.columns:
            remaining_ven_missing[col] = ""
   
    return strict_matches, remaining_erp_missing, remaining_ven_missing
def levenshtein(a, b):
    """Calculate the Levenshtein distance between two strings."""
    if len(a) < len(b):
        a, b = b, a
    if len(b) == 0:
        return len(a)
    previous_row = range(len(b) + 1)
    for i, c1 in enumerate(a):
        current_row = [i + 1]
        for j, c2 in enumerate(b):
            insertions = previous_row[j + 1] + 1
            deletions = current_row[j] + 1
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row
    return previous_row[-1]
def tier2_match(erp_missing, ven_missing):
    """Match unmatched invoices with fuzzy logic and Levenshtein distance, using due date as alternative condition, ignoring amount for matching."""
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), pd.DataFrame(), set(), set(), erp_missing.copy(), ven_missing.copy()
   
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp", "Due Date": "due_date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven", "Due Date": "due_date_ven"}).copy()
   
    # Ensure date and due date columns exist
    for col in ["date_erp", "due_date_erp"]:
        if col not in e_df.columns:
            e_df[col] = ""
    for col in ["date_ven", "due_date_ven"]:
        if col not in v_df.columns:
            v_df[col] = ""
   
    e_df["date_norm"] = e_df["date_erp"].apply(normalize_date)
    v_df["date_norm"] = v_df["date_ven"].apply(normalize_date)
    e_df["due_date_norm"] = e_df["due_date_erp"].apply(normalize_date)
    v_df["due_date_norm"] = v_df["due_date_ven"].apply(normalize_date)
   
    matches, amount_diff_matches, used_e, used_v = [], [], set(), set()
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e: continue
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0)), 2)
        e_date = e.get("date_norm", "")
        e_due = e.get("due_date_norm", "")
        e_code = clean_invoice_code_strict(e_inv)
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0)), 2)
            v_date = v.get("date_norm", "")
            v_due = v.get("due_date_norm", "")
            v_code = clean_invoice_code_strict(v_inv)
            diff = abs(e_amt - v_amt)
            dist = levenshtein(e_code, v_code)
            max_len = max(len(e_code), len(v_code))
            sim = 1 - (dist / max_len) if max_len > 0 else 0
            date_match = (e_date == v_date) or (e_due == v_due)
            if sim >= 0.8 and date_match:
                match_info = {
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Fuzzy Score": round(sim, 2),
                    "Date": e_date or v_date or "N/A",
                    "Due Date": e_due or v_due or "N/A",
                    "Match Type": "Tier-2"
                }
                if diff < 0.05:
                    matches.append(match_info)
                else:
                    match_info["Match Type"] = "Tier-2 Amount Diff"
                    amount_diff_matches.append(match_info)
                used_e.add(e_idx)
                used_v.add(v_idx)
                break
   
    tier2_matches = pd.DataFrame(matches)
    tier2_amount_diff = pd.DataFrame(amount_diff_matches)
   
    erp_columns = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in e_df.columns else []) + (["due_date_erp"] if "due_date_erp" in e_df.columns else [])
    ven_columns = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in v_df.columns else []) + (["due_date_ven"] if "due_date_ven" in v_df.columns else [])
    remaining_erp_missing = e_df[~e_df.index.isin(used_e)][erp_columns].rename(
        columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date", "due_date_erp": "Due Date"}
    )
    remaining_ven_missing = v_df[~v_df.index.isin(used_v)][ven_columns].rename(
        columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date", "due_date_ven": "Due Date"}
    )
   
    return tier2_matches, tier2_amount_diff, used_e, used_v, remaining_erp_missing, remaining_ven_missing
# ======================================
# MAIN STREAMLIT UI
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])
if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")
        with st.spinner("🔍 Analyzing and reconciling invoices..."):
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
            erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)
            strict_matches, final_erp_missing, final_ven_missing = strict_match_unmatched(erp_missing, ven_missing)
            tier2_matches, tier2_amount_diff, _, _, _, _ = tier2_match(final_erp_missing, final_ven_missing)
            # Update final missing after strict and tier2 matches
            erp_missing = final_erp_missing
            ven_missing = final_ven_missing
        st.success("✅ Reconciliation complete!")
       
        # ======================================
        # EXECUTIVE SUMMARY WITH TOTALS
        # ======================================
        st.markdown("## 📈 Executive Summary")
        summary_table = create_summary_table_with_totals(matched, pd.concat([tier2_matches, strict_matches]), tier2_amount_diff, erp_missing, ven_missing)
        st.dataframe(
            style_summary_table(summary_table),
            use_container_width=True,
            hide_index=True
        )
       
        # ======================================
        # ENHANCED METRICS
        # ======================================
        col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
        perfect_count = len(matched[matched['Status'] == 'Perfect Match']) if not matched.empty else 0
        diff_count = len(matched[matched['Status'] == 'Difference Match']) if not matched.empty else 0
        strict_count = len(strict_matches) if not strict_matches.empty else 0
        tier2_count = len(tier2_matches) if not tier2_matches.empty else 0
        tier2_amount_diff_count = len(tier2_amount_diff) if not tier2_amount_diff.empty else 0
        erp_unmatched = len(erp_missing)
        ven_unmatched = len(ven_missing)
        total_reconciled = perfect_count + diff_count + strict_count + tier2_count + tier2_amount_diff_count
        with col1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("🎯 Perfect Matches", perfect_count)
            st.markdown('</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
            st.metric("⚠️ Differences", diff_count)
            st.markdown('</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
            st.metric("🔍 Strict Matches", strict_count)
            st.markdown('</div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
            st.metric("🔍 Tier-2 Matches", tier2_count)
            st.markdown('</div>', unsafe_allow_html=True)
        with col5:
            st.markdown('<div class="metric-container tier2-amount-diff">', unsafe_allow_html=True)
            st.metric("💥 Tier-2 Amount Diff", tier2_amount_diff_count)
            st.markdown('</div>', unsafe_allow_html=True)
        with col6:
            st.markdown('<div class="metric-container payment-match">', unsafe_allow_html=True)
            st.metric("✅ Total Reconciled", total_reconciled)
            st.markdown('</div>', unsafe_allow_html=True)
        with col7:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("❌ ERP Unmatched", erp_unmatched)
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("---")
       
        # ======================================
        # MATCHED INVOICES TABLE
        # ======================================
        st.subheader("✅ MATCHED INVOICES WITH DIFFERENCES")
        if not matched.empty:
            matched_display = matched[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Status']].copy()
            total_row = pd.DataFrame({
                'ERP Invoice': ['TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [matched_display['ERP Amount'].sum()],
                'Vendor Amount': [matched_display['Vendor Amount'].sum()],
                'Difference': [abs(matched_display['ERP Amount'].sum() - matched_display['Vendor Amount'].sum())],
                'Status': [f"TOTAL ({len(matched_display)} MATCHES)"]
            })
            matched_with_totals = pd.concat([matched_display, total_row], ignore_index=True)
            st.dataframe(matched_with_totals, use_container_width=True, height=400)
        else:
            st.info("❌ No Tier-1 matches/differences found.")
       
        # ======================================
        # STRICT MATCHES TABLE
        # ======================================
        if not strict_matches.empty:
            st.subheader("🔍 Strict Matches (Unmatched)")
            strict_display = strict_matches[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Match Type']].copy()
            total_row_strict = pd.DataFrame({
                'ERP Invoice': ['STRICT TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [strict_display['ERP Amount'].sum()],
                'Vendor Amount': [strict_display['Vendor Amount'].sum()],
                'Difference': [abs(strict_display['ERP Amount'].sum() - strict_display['Vendor Amount'].sum())],
                'Match Type': [f"TOTAL ({len(strict_display)} MATCHES)"]
            })
            strict_with_totals = pd.concat([strict_display, total_row_strict], ignore_index=True)
            st.dataframe(strict_with_totals, use_container_width=True, height=400)
       
        # ======================================
        # UNMATCHED INVOICES
        # ======================================
        st.subheader("❌ UNMATCHED INVOICES")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**🔴 Missing in ERP (Vendor Only)**")
            if not ven_missing.empty:
                ven_display = ven_missing.copy()
                total_row_ven = pd.DataFrame({
                    'Invoice': ['TOTAL UNMATCHED'],
                    'Amount': [ven_missing['Amount'].sum()],
                    'Date': [f"{len(ven_missing)} INVOICES"]
                })
                ven_with_total = pd.concat([ven_display, total_row_ven], ignore_index=True)
                st.dataframe(ven_with_total, use_container_width=True)
                st.error(f"**{len(ven_missing)} UNMATCHED | €{ven_missing['Amount'].sum():,.2f}**")
            else:
                st.success("✅ No unmatched vendor invoices!")
        with col2:
            st.markdown("**🔴 Missing in Vendor (ERP Only)**")
            if not erp_missing.empty:
                erp_display = erp_missing.copy()
                total_row_erp = pd.DataFrame({
                    'Invoice': ['TOTAL UNMATCHED'],
                    'Amount': [erp_missing['Amount'].sum()],
                    'Date': [f"{len(erp_missing)} INVOICES"]
                })
                erp_with_total = pd.concat([erp_display, total_row_erp], ignore_index=True)
                st.dataframe(erp_with_total, use_container_width=True)
                st.error(f"**{len(erp_missing)} UNMATCHED | €{erp_missing['Amount'].sum():,.2f}**")
            else:
                st.success("✅ No unmatched ERP invoices!")
       
        # ======================================
        # TIER-2 MATCHES
        # ======================================
        if not tier2_matches.empty:
            st.subheader("🔍 Tier-2 Fuzzy Matches")
            tier2_display = tier2_matches[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Fuzzy Score']].copy()
            total_row_tier2 = pd.DataFrame({
                'ERP Invoice': ['TIER-2 TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [tier2_display['ERP Amount'].sum()],
                'Vendor Amount': [tier2_display['Vendor Amount'].sum()],
                'Difference': [abs(tier2_display['ERP Amount'].sum() - tier2_display['Vendor Amount'].sum())],
                'Fuzzy Score': [f"{len(tier2_display)} MATCHES"]
            })
            tier2_with_total = pd.concat([tier2_display, total_row_tier2], ignore_index=True)
            st.dataframe(tier2_with_total, use_container_width=True)
       
        # ======================================
        # TIER-2 AMOUNT DIFF MATCHES
        # ======================================
        if not tier2_amount_diff.empty:
            st.subheader("💥 Tier-2 Fuzzy Matches with Amount Differences")
            tier2_diff_display = tier2_amount_diff[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Fuzzy Score']].copy()
            total_row_diff = pd.DataFrame({
                'ERP Invoice': ['TIER-2 DIFF TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [tier2_diff_display['ERP Amount'].sum()],
                'Vendor Amount': [tier2_diff_display['Vendor Amount'].sum()],
                'Difference': [abs(tier2_diff_display['ERP Amount'].sum() - tier2_diff_display['Vendor Amount'].sum())],
                'Fuzzy Score': [f"{len(tier2_diff_display)} MATCHES"]
            })
            tier2_diff_with_total = pd.concat([tier2_diff_display, total_row_diff], ignore_index=True)
            st.dataframe(tier2_diff_with_total, use_container_width=True)
       
        # ======================================
        # DOWNLOAD
        # ======================================
        st.markdown("### 📥 Download Full Report")
        excel_output = export_reconciliation_excel(
            matched,
            erp_missing,
            ven_missing,
            matched_pay,
            pd.concat([tier2_matches, strict_matches]),
            tier2_amount_diff
        )
        st.download_button(
            "💾 Download Excel Report",
            data=excel_output,
            file_name="ReconRaptor_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
   
    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        st.info("Please check that your Excel files have the expected columns (invoice, amount, date, etc.)")
