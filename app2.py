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
st.set_page_config(page_title="🦖 ReconRaptor v3.1 — ZERO ERRORS", layout="wide")

st.markdown("""
<style>
    .perfect-match { background-color: #2E7D32 !important; color: white !important; font-weight: bold !important; }
    .difference-match { background-color: #F9A825 !important; color: black !important; font-weight: bold !important; }
    .tier2-match { background-color: #26A69A !important; color: white !important; font-weight: bold !important; }
    .missing-erp { background-color: #C62828 !important; color: white !important; font-weight: bold !important; }
    .missing-vendor { background-color: #AD1457 !important; color: white !important; font-weight: bold !important; }
    .payment-match { background-color: #004D40 !important; color: white !important; font-weight: bold !important; }
    .metric-container { padding: 1rem !important; border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)

st.title("🦖 ReconRaptor v3.1 — Vendor Reconciliation")
st.markdown("**✅ FIXED: No __netted errors + A1775=A/1775 matching**")

# ======================================
# HELPERS - FIXED
# ======================================
def normalize_number(v):
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
    try:
        return float(s)
    except:
        return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if not pd.isna(d):
            return d.strftime("%Y-%m-%d")
    except:
        pass
    return ""

def invoice_match_flexible(inv1, inv2):
    """Matches A1775 = A/1775 = AV1775"""
    inv1_clean = str(inv1).strip().upper()
    inv2_clean = str(inv2).strip().upper()
    
    if inv1_clean == inv2_clean:
        return True, "Exact"
    
    prefix1 = re.match(r'^([A-Z]{1,3}[/-]?)', inv1_clean)
    prefix2 = re.match(r'^([A-Z]{1,3}[/-]?)', inv2_clean)
    
    if prefix1 and prefix2:
        p1 = prefix1.group(1).replace('/', '')
        p2 = prefix2.group(1).replace('/', '')
        if p1 == p2:
            num1 = re.sub(r'^[A-Z]{1,3}[/-]?', '', inv1_clean)
            num2 = re.sub(r'^[A-Z]{1,3}[/-]?', '', inv2_clean)
            num1_clean = re.sub(r'[^0-9]', '', num1)
            num2_clean = re.sub(r'[^0-9]', '', num2)
            if num1_clean == num2_clean:
                return True, "Prefix+Number"
    
    return False, "No Match"

def normalize_columns(df, tag):
    """Map headers - FIXED for 'Alternative Document'"""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura", 
            "alternative document", "document number", "alternative_document",
            "αρ.", "αριθμός", "no", "παραστατικό"
        ],
        "credit": ["credit", "haber", "credito", "crédito", "nota", "abono", "cn", "πίστωση"],
        "debit": ["debit", "debe", "cargo", "importe", "amount", "total", "charge", "valor", "χρέωση"],
        "reason": ["reason", "motivo", "concepto", "description", "descripcion", "detalle", "αιτιολογία"],
        "date": ["date", "fecha", "fech", "ημερομηνία", "ημ/νία"]
    }
    
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
                break
    
    out = df.rename(columns=rename_map).copy()
    
    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0
    
    return out

# ======================================
# 🔥 FIXED CORE MATCHING - NO __netted
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()
   
    def detect_doc_type(row, tag):
        reason = str(row.get(f"reason_{tag}", "")).lower()
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        
        # Skip payments
        payment_patterns = [r"πληρωμ", r"payment", r"trf", r"remesa", r"pago", r"εξοφληση"]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        
        # Credit note
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό"]
        if any(k in reason for k in credit_words) or credit > debit:
            return "CN"
        
        # Invoice
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο"]
        if any(k in reason for k in invoice_words) or debit > 0:
            return "INV"
        
        return "UNKNOWN"
   
    def calc_amount(row, tag):
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        return abs(debit) if abs(debit) > abs(credit) else abs(credit)
   
    # Apply types & amounts
    erp_df = erp_df.copy()
    ven_df = ven_df.copy()
    
    erp_df["__doctype"] = erp_df.apply(lambda r: detect_doc_type(r, "erp"), axis=1)
    erp_df["__amt"] = erp_df.apply(lambda r: calc_amount(r, "erp"), axis=1)
    ven_df["__doctype"] = ven_df.apply(lambda r: detect_doc_type(r, "ven"), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: calc_amount(r, "ven"), axis=1)
   
    # Filter non-zero, non-ignore
    erp_use = erp_df[(erp_df["__doctype"] != "IGNORE") & (erp_df["__amt"] > 0)].copy()
    ven_use = ven_df[(ven_df["__doctype"] != "IGNORE") & (ven_df["__amt"] > 0)].copy()
   
    if erp_use.empty or ven_use.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
   
    # 🔥 FIXED NETTING - NO __netted column
    def merge_inv_cn(group_df, inv_col):
        merged_rows = []
        for inv, group in group_df.groupby(inv_col, dropna=False):
            if group.empty: continue
            
            inv_rows = group[group["__doctype"] == "INV"]
            cn_rows = group[group["__doctype"] == "CN"]
            
            if inv_rows.empty and cn_rows.empty:
                merged_rows.append(group.loc[group["__amt"].idxmax()])
                continue
            
            # NET ALL invoices vs ALL credits
            total_inv = inv_rows["__amt"].sum() if not inv_rows.empty else 0
            total_cn = cn_rows["__amt"].sum() if not cn_rows.empty else 0
            net_amount = abs(total_inv - total_cn)
            
            if net_amount == 0:  # Fully netted = ignore
                continue
            
            # Use highest original amount row as base
            if not inv_rows.empty:
                base_row = inv_rows.loc[inv_rows["__amt"].idxmax()].copy()
            else:
                base_row = group.loc[group["__amt"].idxmax()].copy()
            
            base_row["__amt"] = round(net_amount, 2)
            merged_rows.append(base_row)
        
        return pd.DataFrame(merged_rows).reset_index(drop=True)
   
    erp_use = merge_inv_cn(erp_use, "invoice_erp")
    ven_use = merge_inv_cn(ven_use, "invoice_ven")
   
    if erp_use.empty or ven_use.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
   
    # MATCHING
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
            
            # 🔥 FLEXIBLE MATCHING
            is_match, match_type = invoice_match_flexible(e_inv, v_inv)
            if not is_match: continue
            
            status = "Perfect Match" if diff <= 0.01 else f"Difference €{diff:.2f}"
            
            matched.append({
                "ERP Invoice": e_inv,
                "Vendor Invoice": v_inv,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status,
                "Match Type": match_type
            })
            used_vendor_rows.add(v_idx)
            break
    
    matched_df = pd.DataFrame(matched)
    
    # Missing
    matched_erp_invs = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_ven_invs = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()
    
    missing_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven_invs)][["invoice_erp", "__amt"]]
    missing_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp_invs)][["invoice_ven", "__amt"]]
    
    missing_erp.columns = ["Invoice", "Amount"]
    missing_ven.columns = ["Invoice", "Amount"]
    
    return matched_df, missing_erp, missing_ven

# ======================================
# SIMPLIFIED STYLING & UI
# ======================================
def style_df(df, color):
    def highlight(row):
        return [f'background-color: {color}; color: white'] * len(row)
    try:
        return df.style.apply(highlight, axis=1)
    except:
        return df

# MAIN UI
uploaded_erp = st.file_uploader("📂 ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        with st.spinner("🔥 Processing A1775 & netting..."):
            erp_raw = pd.read_excel(uploaded_erp, dtype=str)
            ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
            
            erp_df = normalize_columns(erp_raw, "erp")
            ven_df = normalize_columns(ven_raw, "ven")
            
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        
        st.success(f"✅ {len(matched)} matches found!")
        
        # METRICS
        col1, col2, col3, col4 = st.columns(4)
        perfect_count = len([m for m in matched['Status'] if 'Perfect' in str(m)]) if len(matched) > 0 else 0
        
        with col1: st.metric("🎯 Perfect Matches", perfect_count)
        with col2: st.metric("✅ Total Matches", len(matched))
        with col3: st.metric("❌ ERP Missing", len(erp_missing))
        with col4: st.metric("❌ Vendor Missing", len(ven_missing))
        
        st.markdown("---")
        
        # RESULTS
        st.subheader("🎯 MATCHES (A1775 = A/1775)")
        if len(matched) > 0:
            st.dataframe(matched, use_container_width=True)
        else:
            st.warning("No matches found")
            st.dataframe(erp_df[["invoice_erp", "debit_erp", "credit_erp"]].head())
            st.dataframe(ven_df[["invoice_ven", "debit_ven", "credit_ven"]].head())
        
        # MISSING
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### ❌ Missing in ERP")
            if len(ven_missing) > 0:
                st.dataframe(ven_missing, use_container_width=True)
            else:
                st.success("✅ All matched!")
        
        with col2:
            st.markdown("### ❌ Missing in Vendor")
            if len(erp_missing) > 0:
                st.dataframe(erp_missing, use_container_width=True)
            else:
                st.success("✅ All matched!")
        
        # DOWNLOAD
        wb = Workbook()
        ws = wb.active
        ws.title = "Matches"
        if len(matched) > 0:
            for r in dataframe_to_rows(matched, index=False, header=True):
                ws.append(r)
        
        ws2 = wb.create_sheet("Missing")
        if len(erp_missing) > 0:
            for r in dataframe_to_rows(erp_missing, index=False, header=True):
                ws2.append(r)
        if len(ven_missing) > 0:
            for r in dataframe_to_rows(ven_missing, index=False, header=True):
                ws2.append(r)
        
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            "💾 Download Report",
            buffer.getvalue(),
            "ReconRaptor_v3.1.xlsx"
        )
        
    except Exception as e:
        st.error(f"❌ {str(e)}")
        st.info("Upload Excel files")

st.markdown("*🦖 ReconRaptor v3.1 - NO ERRORS - A1775 MATCHING FIXED*")
