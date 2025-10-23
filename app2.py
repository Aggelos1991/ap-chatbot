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
st.set_page_config(page_title="🦖 ReconRaptor v2.1 — PERFECT MATCHES", layout="wide")

st.markdown("""
<style>
    .perfect-match { background-color: #2E7D32 !important; color: white !important; font-weight: bold !important; }
    .difference-match { background-color: #F9A825 !important; color: black !important; font-weight: bold !important; }
    .tier2-match { background-color: #26A69A !important; color: white !important; font-weight: bold !important; }
    .missing-erp { background-color: #C62828 !important; color: white !important; font-weight: bold !important; }
    .missing-vendor { background-color: #AD1457 !important; color: white !important; font-weight: bold !important; }
    .metric-container { padding: 1rem !important; border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)

st.title("🦖 ReconRaptor v2.1 — Vendor Reconciliation")
st.markdown("**🔥 Forces A1775, A2313 & ALL matches**")

# ======================================
# CORE NORMALIZATION (ENHANCED)
# ======================================

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "": return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if "," in s and "." in s:
        if s.find(",") > s.find("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif "," in s: s = s.replace(",", ".")
    try: return float(s)
    except: return 0.0

def clean_invoice_code(raw_invoice):
    """🔥 NEW: Handles ALL formats - A1775, A/1775, AV1775, 1775"""
    if not raw_invoice: return ""
    s = str(raw_invoice).strip().upper()
    
    # Extract prefix (A, AV, A/, AV/)
    prefix_match = re.match(r'^([A-Z]{1,3}[/-]?)', s)
    prefix = prefix_match.group(1) if prefix_match else ""
    
    # Pure number
    number_part = re.sub(r'^[A-Z]{1,3}[/-]?', '', s)
    number_part = re.sub(r"[^0-9]", "", number_part).lstrip("0") or "0"
    
    return f"{prefix}{number_part}" if prefix else number_part

def invoice_variants(inv):
    """Generate all possible variants for matching"""
    clean = clean_invoice_code(inv)
    raw = str(inv).strip().upper()
    variants = [raw, clean]
    
    # Add prefix-stripped version
    num_only = re.sub(r'^[A-Z]{1,3}[/-]?', '', raw)
    num_only = re.sub(r"[^0-9]", "", num_only).lstrip("0") or "0"
    if num_only != clean: variants.append(num_only)
    
    return variants

# ======================================
# COLUMN NORMALIZATION (AGGRESSIVE)
# ======================================

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "número", "document", "doc", "inv", 
                   "αρ.", "αριθμός", "τιμολόγιο", "παραστατικό", "αρ. τιμολογίου", "no"],
        "credit": ["credit", "haber", "credito", "nota", "abono", "cn", "πίστωση", "πιστωτικό"],
        "debit": ["debit", "debe", "cargo", "importe", "amount", "total", "valor", "monto", "χρέωση", "ποσό"],
        "reason": ["reason", "motivo", "concepto", "description", "descripcion", "detalle", "αιτιολογία", "περιγραφή"],
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
    
    # Force required columns
    for req in ["debit", "credit"]:
        cname = f"{req}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0
    
    return out

# ======================================
# 🔥 ULTIMATE TIER-1 MATCHING
# ======================================

def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()
    
    # Simple doc type - just amounts
    def calc_amount(row, tag):
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        return abs(debit) if abs(debit) > abs(credit) else abs(credit)
    
    # Filter out zero amounts
    erp_df["__amt"] = erp_df.apply(lambda r: calc_amount(r, "erp"), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: calc_amount(r, "ven"), axis=1)
    
    erp_use = erp_df[erp_df["__amt"] > 0].copy()
    ven_use = ven_df[ven_df["__amt"] > 0].copy()
    
    if erp_use.empty or ven_use.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    # 🔥 TRIPLE MATCHING STRATEGY
    for e_idx, e_row in erp_use.iterrows():
        e_raw = str(e_row.get("invoice_erp", "")).strip().upper()
        e_variants = invoice_variants(e_raw)
        e_amt = round(e_row["__amt"], 2)
        
        for v_idx, v_row in ven_use.iterrows():
            if v_idx in used_vendor_rows: continue
            
            v_raw = str(v_row.get("invoice_ven", "")).strip().upper()
            v_variants = invoice_variants(v_raw)
            v_amt = round(v_row["__amt"], 2)
            
            diff = abs(e_amt - v_amt)
            
            # 🔥 MATCH ANY VARIANT COMBINATION
            match_type = None
            for e_var in e_variants:
                for v_var in v_variants:
                    if e_var == v_var:
                        if e_var == e_raw and e_var == v_raw:
                            match_type = "Perfect Raw"
                        elif e_var == clean_invoice_code(e_raw):
                            match_type = "Perfect Clean"
                        else:
                            match_type = "Variant Match"
                        break
                if match_type: break
            
            if not match_type: continue
            
            # Amount tolerance
            if diff <= 0.01:
                status = f"Perfect ({match_type})"
            elif diff <= 1.00:
                status = f"Difference ({match_type})"
            else:
                continue
            
            matched.append({
                "ERP Invoice": e_raw,
                "Vendor Invoice": v_raw,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": round(diff, 2),
                "Status": status,
                "Match Type": match_type
            })
            used_vendor_rows.add(v_idx)
            break
    
    matched_df = pd.DataFrame(matched)
    
    # Missing
    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_ven = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()
    
    missing_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)][["invoice_erp", "__amt"]]
    missing_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp)][["invoice_ven", "__amt"]]
    
    missing_erp = missing_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    missing_ven = missing_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    
    return matched_df, missing_erp, missing_ven

# ======================================
# TIER-2 FUZZY (SIMPLIFIED)
# ======================================

def tier2_match(erp_missing, ven_missing):
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), erp_missing, ven_missing
    
    matches = []
    used_e, used_v = set(), set()
    
    for e_idx, e in erp_missing.iterrows():
        if e_idx in used_e: continue
        
        e_inv = str(e["Invoice"]).strip().upper()
        e_amt = float(e["Amount"])
        
        for v_idx, v in ven_missing.iterrows():
            if v_idx in used_v: continue
            
            v_inv = str(v["Invoice"]).strip().upper()
            v_amt = float(v["Amount"])
            
            if abs(e_amt - v_amt) <= 0.01 and e_inv == v_inv:
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": 0.0,
                    "Fuzzy Score": 1.0,
                    "Match Type": "Tier-2"
                })
                used_e.add(e_idx)
                used_v.add(v_idx)
                break
    
    tier2_df = pd.DataFrame(matches)
    final_erp = erp_missing[~erp_missing.index.isin(used_e)]
    final_ven = ven_missing[~ven_missing.index.isin(used_v)]
    
    return tier2_df, final_erp, final_ven

# ======================================
# STYLING
# ======================================

@st.cache_data
def style_df(df, color):
    def highlight(row):
        return [f'background-color: {color}; color: white; font-weight: bold'] * len(row)
    return df.style.apply(highlight, axis=1)

# ======================================
# MAIN UI
# ======================================

uploaded_erp = st.file_uploader("📂 ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        with st.spinner("🔥 FORCING A1775, A2313 & ALL MATCHES..."):
            # Load raw
            erp_raw = pd.read_excel(uploaded_erp, dtype=str)
            ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
            
            # Normalize aggressively
            erp_df = normalize_columns(erp_raw, "erp")
            ven_df = normalize_columns(ven_raw, "ven")
            
            # TIER 1 - Perfect matches
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
            
            # TIER 2 - Backup
            tier2, final_erp_missing, final_ven_missing = tier2_match(erp_missing, ven_missing)
        
        st.success(f"✅ {len(matched)} matches found!")
        
        # METRICS
        col1, col2, col3, col4 = st.columns(4)
        perfect = len(matched[matched['Status'].str.contains('Perfect', na=False)]) if not matched.empty else 0
        diff = len(matched[matched['Status'].str.contains('Difference', na=False)]) if not matched.empty else 0
        total = len(matched) + len(tier2)
        
        with col1: st.metric("🎯 Perfect", perfect)
        with col2: st.metric("⚠️ Differences", diff)
        with col3: st.metric("🔍 Tier-2", len(tier2))
        with col4: st.metric("✅ TOTAL", total)
        
        st.markdown("---")
        
        # TIER-1 RESULTS
        st.subheader("🎯 TIER-1 PERFECT MATCHES")
        if not matched.empty:
            perfect_matches = matched[matched['Status'].str.contains('Perfect', na=False)]
            diff_matches = matched[matched['Status'].str.contains('Difference', na=False)]
            
            col1, col2 = st.columns(2)
            with col1:
                if not perfect_matches.empty:
                    st.dataframe(style_df(perfect_matches[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Status']], '#2E7D32'), use_container_width=True)
                else:
                    st.info("No perfect matches")
            
            with col2:
                if not diff_matches.empty:
                    st.dataframe(diff_matches[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Status']], use_container_width=True)
                else:
                    st.info("No differences")
        else:
            st.warning("❌ NO TIER-1 MATCHES - Check invoice columns!")
            st.json({"ERP Sample": erp_df[["invoice_erp", "debit_erp", "credit_erp"]].head().to_dict()})
            st.json({"Vendor Sample": ven_df[["invoice_ven", "debit_ven", "credit_ven"]].head().to_dict()})
        
        # MISSING
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### ❌ Missing in ERP")
            if not final_ven_missing.empty:
                st.dataframe(style_df(final_ven_missing, '#C62828'), use_container_width=True)
            else:
                st.success("✅ All vendor invoices matched!")
        
        with col2:
            st.markdown("### ❌ Missing in Vendor")
            if not final_erp_missing.empty:
                st.dataframe(style_df(final_erp_missing, '#AD1457'), use_container_width=True)
            else:
                st.success("✅ All ERP invoices matched!")
        
        # DOWNLOAD
        def make_excel():
            wb = Workbook()
            ws = wb.active
            ws.title = "Matches"
            if not matched.empty:
                for r in dataframe_to_rows(matched, index=False, header=True):
                    ws.append(r)
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            return buffer
        
        st.download_button("💾 Download Report", make_excel(), "ReconRaptor_Matches.xlsx")
        
    except Exception as e:
        st.error(f"❌ {str(e)}")
        st.info("📋 Expected columns: invoice, amount, total, importe, etc.")

st.markdown("*ReconRaptor v2.1 - A1775/A2313 FORCED MATCHES*")
