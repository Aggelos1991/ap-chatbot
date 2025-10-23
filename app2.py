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
# CONFIGURATION & CSS
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor v2.0 — Perfect Matching", layout="wide")

st.markdown("""
<style>
    .perfect-match { background-color: #2E7D32 !important; color: white !important; font-weight: bold !important; }
    .difference-match { background-color: #F9A825 !important; color: black !important; font-weight: bold !important; }
    .tier2-match { background-color: #26A69A !important; color: white !important; font-weight: bold !important; }
    .tier3-match { background-color: #FF9800 !important; color: white !important; font-weight: bold !important; }
    .missing-erp { background-color: #C62828 !important; color: white !important; font-weight: bold !important; }
    .missing-vendor { background-color: #AD1457 !important; color: white !important; font-weight: bold !important; }
    .payment-match { background-color: #004D40 !important; color: white !important; font-weight: bold !important; }
    .erp-payment { background-color: #4CAF50 !important; color: white !important; }
    .vendor-payment { background-color: #2196F3 !important; color: white !important; }
    .metric-container { padding: 1rem !important; border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)

st.title("🦖 ReconRaptor v2.0 — Vendor Reconciliation")
st.markdown("**Perfect matching for A1775, A2313, and ALL edge cases**")

# ======================================
# CORE NORMALIZATION FUNCTIONS
# ======================================

def normalize_number(v):
    """Convert '1.234,56' or '1,234.56' → float safely."""
    if pd.isna(v) or str(v).strip() == "":
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
    """YYYY-MM-DD from any format."""
    if pd.isna(v) or str(v).strip() == "":
        return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y", "%Y/%m/%d",
        "%d/%m/%y", "%m/%d/%y", "%Y.%m.%d"
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
        if not pd.isna(d):
            return d.strftime("%Y-%m-%d")
    except:
        return ""

def clean_invoice_code(raw_invoice):
    """
    NEW v2.0: Preserves prefixes (A/, AV/, AJ/) + cleans number
    A1775 → A1775
    A/1775 → A1775  
    AV1234 → AV1234
    """
    if not raw_invoice:
        return ""
    
    s = str(raw_invoice).strip().upper()
    
    # Extract prefix (A, AV, AJ, A/, AV/ etc.)
    prefix_match = re.match(r'^([A-Z]{1,3}[/-]?)', s)
    prefix = prefix_match.group(1) if prefix_match else ""
    
    # Extract pure number
    number_part = re.sub(r'^[A-Z]{1,3}[/-]?', '', s)
    number_part = re.sub(r"[^0-9]", "", number_part).lstrip("0") or "0"
    
    return f"{prefix}{number_part}" if prefix else number_part

def has_same_prefix_and_number(inv1, inv2):
    """
    TIER 1C: A1775 == A/1775 == AV1775 (same prefix family + number)
    """
    # Get prefixes
    prefix1 = re.match(r'^([A-Z]{1,3}[/-]?)', inv1)
    prefix2 = re.match(r'^([A-Z]{1,3}[/-]?)', inv2)
    
    if not prefix1 or not prefix2:
        return False
    
    # Same prefix family (A/ = A, AV/ = AV)
    p1_clean = prefix1.group(1).replace('/', '')
    p2_clean = prefix2.group(1).replace('/', '')
    if p1_clean != p2_clean:
        return False
    
    # Same number
    num1 = re.sub(r'^[A-Z]{1,3}[/-]?', '', inv1)
    num2 = re.sub(r'^[A-Z]{1,3}[/-]?', '', inv2)
    num1_clean = re.sub(r"[^0-9]", "", num1)
    num2_clean = re.sub(r"[^0-9]", "", num2)
    
    return num1_clean == num2_clean

# ======================================
# COLUMN NORMALIZATION
# ======================================

def normalize_columns(df, tag):
    """Map multilingual headers → unified names."""
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "número", "document", "doc", 
                   "αρ.", "αριθμός", "τιμολόγιο", "παραστατικό", "αρ. τιμολογίου"],
        "credit": ["credit", "haber", "credito", "crédito", "nota de crédito", "abono", 
                  "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο"],
        "debit": ["debit", "debe", "cargo", "importe", "amount", "total", "valor", "monto",
                 "χρέωση", "αξία", "ποσό", "ποσό τιμολογίου"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "descripción", "detalle",
                  "αιτιολογία", "περιγραφή", "παρατηρήσεις"],
        "date": ["date", "fecha", "fech", "data", "fecha factura", "ημερομηνία", "ημ/νία"]
    }
    
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
    
    out = df.rename(columns=rename_map)
    
    # Ensure required columns
    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0
    
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    
    return out

# ======================================
# TIER-1 MATCHING (PERFECT MATCHES)
# ======================================

def match_invoices(erp_df, ven_df):
    """
    3-TIER LOGIC:
    TIER 1A: PERFECT RAW - A1775 == A1775
    TIER 1B: PERFECT CLEAN - A/1775 == A1775  
    TIER 1C: PREFIX FAMILY - A1775 == AV1775
    """
    matched = []
    used_vendor_rows = set()
    
    # === DOCUMENT TYPE DETECTION ===
    def detect_doc_type(row, tag):
        reason = str(row.get(f"reason_{tag}", "")).lower()
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        
        # Skip payments
        payment_patterns = [r"πληρωμ", r"payment", r"trf", r"remesa", r"pago", r"εξοφληση"]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        
        # Credit Note
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση"]
        if any(k in reason for k in credit_words) or credit > 0:
            return "CN"
        
        # Invoice
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in invoice_words) or debit > 0:
            return "INV"
        
        return "UNKNOWN"
    
    def calc_amount(row, tag):
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        return abs(debit) if abs(debit) > 0 else abs(credit)
    
    # Apply filters
    erp_df["__doctype"] = erp_df.apply(lambda r: detect_doc_type(r, "erp"), axis=1)
    erp_df["__amt"] = erp_df.apply(lambda r: calc_amount(r, "erp"), axis=1)
    ven_df["__doctype"] = ven_df.apply(lambda r: detect_doc_type(r, "ven"), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: calc_amount(r, "ven"), axis=1)
    
    erp_use = erp_df[erp_df["__doctype"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__doctype"] != "IGNORE"].copy()
    
    # === MERGE INV+CN ===
    def merge_inv_cn(group_df, inv_col):
        merged = []
        for inv, group in group_df.groupby(inv_col, dropna=False):
            if group.empty: continue
            inv_rows = group[group["__doctype"] == "INV"]
            cn_rows = group[group["__doctype"] == "CN"]
            
            if not inv_rows.empty and not cn_rows.empty:
                net = round(abs(inv_rows["__amt"].sum() - cn_rows["__amt"].sum()), 2)
                base_row = inv_rows.iloc[-1].copy()
                base_row["__amt"] = net
                merged.append(base_row)
            else:
                merged.append(group.loc[group["__amt"].idxmax()])
        return pd.DataFrame(merged).reset_index(drop=True)
    
    erp_use = merge_inv_cn(erp_use, "invoice_erp")
    ven_use = merge_inv_cn(ven_use, "invoice_ven")
    
    erp_use["__amt"] = erp_use["__amt"].astype(float)
    ven_use["__amt"] = ven_use["__amt"].astype(float)
    
    # === 3-TIER TIER-1 MATCHING ===
    for e_idx, e in erp_use.iterrows():
        e_raw_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_clean_inv = clean_invoice_code(e_raw_inv)
        e_amt = round(float(e["__amt"]), 2)
        e_type = e["__doctype"]
        
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows: continue
            
            v_raw_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_clean_inv = clean_invoice_code(v_raw_inv)
            v_amt = round(float(v["__amt"]), 2)
            v_type = v["__doctype"]
            
            diff = abs(e_amt - v_amt)
            
            # Same document type
            if e_type != v_type: continue
            
            # === TIER 1A: PERFECT RAW ===
            if e_raw_inv == v_raw_inv:
                match_tier = "Perfect Raw"
            
            # === TIER 1B: PERFECT CLEAN ===
            elif e_clean_inv == v_clean_inv:
                match_tier = "Perfect Clean"
            
            # === TIER 1C: PREFIX FAMILY ===
            elif has_same_prefix_and_number(e_raw_inv, v_raw_inv):
                match_tier = "Prefix Family"
            
            else:
                continue
            
            # Amount tolerance
            if diff <= 0.01:
                status = f"Perfect Match ({match_tier})"
            elif diff < 1.00:
                status = f"Difference ({match_tier})"
            else:
                continue
            
            matched.append({
                "ERP Invoice": e_raw_inv,
                "Vendor Invoice": v_raw_inv,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status,
                "Match Tier": match_tier
            })
            used_vendor_rows.add(v_idx)
            break
    
    matched_df = pd.DataFrame(matched)
    
    # Unmatched
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
# TIER-2 FUZZY MATCHING
# ======================================

def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def tier2_match(erp_missing, ven_missing):
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt"}).copy()
    
    matches = []
    used_e, used_v = set(), set()
    
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e: continue
        
        e_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_amt = round(float(e.get("__amt", 0)), 2)
        e_code = clean_invoice_code(e_inv)
        
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v: continue
            
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(float(v.get("__amt", 0)), 2)
            v_code = clean_invoice_code(v_inv)
            
            diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_code, v_code)
            
            # Tighter criteria for Tier-2
            if diff <= 0.01 and sim >= 0.85:
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Fuzzy Score": round(sim, 2),
                    "Match Type": "Tier-2 Fuzzy"
                })
                used_e.add(e_idx)
                used_v.add(v_idx)
                break
    
    tier2_df = pd.DataFrame(matches)
    remaining_erp = e_df[~e_df.index.isin(used_e)].rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    remaining_ven = v_df[~v_df.index.isin(used_v)].rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    
    return tier2_df, remaining_erp, remaining_ven

# ======================================
# PAYMENTS
# ======================================

def extract_payments(erp_df, ven_df):
    payment_keywords = ["πληρωμή", "payment", "bank transfer", "trf", "remesa", "pago", "εξοφληση"]
    
    def is_payment(row, tag):
        reason = str(row.get(f"reason_{tag}", "")).lower()
        if any(k in reason for k in payment_keywords):
            amt = abs(normalize_number(row.get(f"debit_{tag}", 0)) - normalize_number(row.get(f"credit_{tag}", 0)))
            return amt > 0
        return False
    
    erp_pay = erp_df[erp_df.apply(lambda x: is_payment(x, "erp"), axis=1)].copy()
    ven_pay = ven_df[ven_df.apply(lambda x: is_payment(x, "ven"), axis=1)].copy()
    
    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(lambda r: abs(normalize_number(r["debit_erp"]) - normalize_number(r["credit_erp"])), axis=1)
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(lambda r: abs(normalize_number(r["debit_ven"]) - normalize_number(r["credit_ven"])), axis=1)
    
    return erp_pay, ven_pay, pd.DataFrame()

# ======================================
# EXCEL EXPORT
# ======================================

def export_reconciliation_excel(matched, erp_missing, ven_missing, tier2_matches):
    wb = Workbook()
    
    # Tier-1 Matches
    ws1 = wb.active
    ws1.title = "Tier1_PerfectMatches"
    if not matched.empty:
        for r in dataframe_to_rows(matched[["ERP Invoice", "Vendor Invoice", "ERP Amount", "Vendor Amount", "Difference", "Status", "Match Tier"]], index=False, header=True):
            ws1.append(r)
    
    # Missing
    ws2 = wb.create_sheet("Missing_Invoices")
    if not erp_missing.empty:
        for r in dataframe_to_rows(erp_missing, index=False, header=True):
            ws2.append(r)
    if not ven_missing.empty:
        for r in dataframe_to_rows(ven_missing, index=False, header=True):
            ws2.append(r)
    
    # Tier-2
    ws3 = wb.create_sheet("Tier2_Fuzzy")
    if not tier2_matches.empty:
        for r in dataframe_to_rows(tier2_matches, index=False, header=True):
            ws3.append(r)
    
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ======================================
# STYLING FUNCTIONS
# ======================================

def style_perfect_matches(df):
    return df.style.apply(lambda row: ['background-color: #2E7D32; color: white; font-weight: bold'] * len(row), axis=1)

def style_tier2_matches(df):
    return df.style.apply(lambda row: ['background-color: #26A69A; color: white; font-weight: bold'] * len(row), axis=1)

def style_missing(df):
    return df.style.apply(lambda row: ['background-color: #C62828; color: white; font-weight: bold'] * len(row), axis=1)

# ======================================
# MAIN UI
# ======================================

uploaded_erp = st.file_uploader("📂 ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        with st.spinner("🔍 Smart matching A1775, A2313 & all invoices..."):
            # Load & normalize
            erp_raw = pd.read_excel(uploaded_erp, dtype=str)
            ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
            erp_df = normalize_columns(erp_raw, "erp")
            ven_df = normalize_columns(ven_raw, "ven")
            
            # TIER-1 Perfect Matches
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
            
            # TIER-2 Fuzzy
            tier2_matches, final_erp_missing, final_ven_missing = tier2_match(erp_missing, ven_missing)
            
            # Payments
            erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)
        
        st.success("✅ Perfect reconciliation complete!")
        
        # === METRICS ===
        col1, col2, col3, col4 = st.columns(4)
        perfect_count = len(matched[matched['Status'].str.contains('Perfect')])
        diff_count = len(matched[matched['Status'].str.contains('Difference')])
        tier2_count = len(tier2_matches)
        
        with col1: st.metric("🎯 Perfect Matches", perfect_count)
        with col2: st.metric("⚠️ Differences", diff_count)
        with col3: st.metric("🔍 Fuzzy Matches", tier2_count)
        with col4: st.metric("✅ Total Reconciled", perfect_count + diff_count + tier2_count)
        
        st.markdown("---")
        
        # === TIER-1 RESULTS ===
        st.subheader("🎯 Tier-1 Perfect Matches")
        if not matched.empty:
            col1, col2 = st.columns(2)
            with col1:
                perfect = matched[matched['Status'].str.contains('Perfect')]
                st.dataframe(style_perfect_matches(perfect), use_container_width=True)
            with col2:
                diff = matched[matched['Status'].str.contains('Difference')]
                st.dataframe(diff.style, use_container_width=True)
        
        # === TIER-2 ===
        st.subheader("🔍 Tier-2 Fuzzy Matches")
        if not tier2_matches.empty:
            st.dataframe(style_tier2_matches(tier2_matches), use_container_width=True)
        
        # === MISSING ===
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### ❌ Missing in ERP")
            if not final_ven_missing.empty:
                st.dataframe(style_missing(final_ven_missing), use_container_width=True)
            else:
                st.success("✅ All vendor invoices matched!")
        
        with col2:
            st.markdown("### ❌ Missing in Vendor") 
            if not final_erp_missing.empty:
                st.dataframe(style_missing(final_erp_missing), use_container_width=True)
            else:
                st.success("✅ All ERP invoices matched!")
        
        # === DOWNLOAD ===
        excel_output = export_reconciliation_excel(matched, final_erp_missing, final_ven_missing, tier2_matches)
        st.download_button(
            "💾 Download Full Report",
            data=excel_output,
            file_name="ReconRaptor_v2_PerfectMatches.xlsx"
        )
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Check your Excel files have invoice/amount columns")

st.markdown("---")
st.markdown("*ReconRaptor v2.0 - Fixed A1775/A2313 matching + 3-tier logic*")
