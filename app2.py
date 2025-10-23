import streamlit as st
import pandas as pd
import re
from io import BytesIO
from difflib import SequenceMatcher

# ======================================
# CONFIGURATION
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")

st.markdown("""
<style>
    .perfect-match { background-color: #2E7D32 !important; color: white !important; font-weight: bold !important; }
    .difference-match { background-color: #F9A825 !important; color: black !important; font-weight: bold !important; }
    .tier2-strict { background-color: #26A69A !important; color: white !important; font-weight: bold !important; }
    .tier2-relaxed { background-color: #FFCA28 !important; color: black !important; font-weight: bold !important; }
    .missing-erp { background-color: #C62828 !important; color: white !important; font-weight: bold !important; }
    .missing-vendor { background-color: #AD1457 !important; color: white !important; font-weight: bold !important; }
    .metric-container { padding: 1rem !important; border-radius: 10px !important; }
    .total-row { background: linear-gradient(90deg, #667eea 0%, #764ba2 100%) !important; color: white !important; font-weight: bold !important; font-size: 14px !important; }
</style>
""", unsafe_allow_html=True)

st.title("🦖 ReconRaptor — Vendor Reconciliation")

# ======================================
# HELPERS
# ======================================
def normalize_number(v):
    if v is None or str(v).strip() == "": return 0.0
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
    if pd.isna(v) or str(v).strip() == "": return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    formats = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y", "%Y/%m/%d", "%d/%m/%y", "%Y-%m-%d"]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d): return d.strftime("%Y-%m-%d")
        except: continue
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if not pd.isna(d): return d.strftime("%Y-%m-%d")
    except: pass
    return ""

def clean_invoice_code(v):
    if not v: return "0"
    s = str(v).strip().lower()
    parts = re.split(r"[-_]", s)
    for p in reversed(parts):
        if re.fullmatch(r"\d{3,}", p) and not re.fullmatch(r"20[0-3]\d", p):
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
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "document", "doc", "ref", "αρ.", "αριθμός", "παραστατικό", "τιμολόγιο"],
        "credit": ["credit", "haber", "credito", "nota", "abono", "πίστωση", "πιστωτικό"],
        "debit": ["debit", "debe", "cargo", "importe", "amount", "total", "valor", "monto", "χρέωση", "αξία", "ποσό"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "description", "αιτιολογία", "περιγραφή"],
        "date": ["date", "fecha", "ημερομηνία", "ημ/νία"]
    }
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
                break
    out = df.rename(columns=rename_map)
    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

# ======================================
# TIER-2 MATCHING FUNCTIONS
# ======================================
def tier2_match_strict(erp_missing, ven_missing):
    """🔒 TIER-2 STRICT: Fuzzy invoice + EXACT amounts (±0.05)"""
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), set(), set(), erp_missing.copy(), ven_missing.copy()
    
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    
    matches, used_e, used_v = [], set(), set()
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e: continue
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0)), 2)
        e_code = clean_invoice_code(e_inv)
        
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0)), 2)
            v_code = clean_invoice_code(v_inv)
            
            diff = abs(e_amt - v_amt)
            sim = SequenceMatcher(None, e_code, v_code).ratio()
            
            # STRICT: High fuzzy + exact amounts
            if diff <= 0.05 and sim >= 0.8:
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Fuzzy Score": f"{sim:.1%}",
                    "Match Type": "🔒 Tier-2 Strict"
                })
                used_e.add(e_idx)
                used_v.add(v_idx)
                break
    
    tier2_matches = pd.DataFrame(matches)
    remaining_erp = e_df[~e_df.index.isin(used_e)]
    remaining_ven = v_df[~v_df.index.isin(used_v)]
    
    return tier2_matches, used_e, used_v, remaining_erp, remaining_ven

def tier2_match_relaxed(erp_missing, ven_missing):
    """🟡 TIER-2 RELAXED: Fuzzy invoice ONLY - NO amount limit"""
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame()
    
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    
    matches = []
    for e_idx, e in e_df.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0)), 2)
        e_code = clean_invoice_code(e_inv)
        
        for v_idx, v in v_df.iterrows():
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0)), 2)
            v_code = clean_invoice_code(v_inv)
            
            sim = SequenceMatcher(None, e_code, v_code).ratio()
            diff = abs(e_amt - v_amt)
            
            # RELAXED: ONLY fuzzy invoice pattern match - ANY amount difference OK
            if sim >= 0.85:  # 85%+ invoice similarity
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Amount Diff": diff,
                    "Fuzzy Score": f"{sim:.1%}",
                    "ERP Date": e.get("date_erp", ""),
                    "Vendor Date": v.get("date_ven", ""),
                    "Match Type": "🟡 Tier-2 Relaxed"
                })
                break  # One match per ERP invoice
    
    return pd.DataFrame(matches)

# ======================================
# OTHER FUNCTIONS
# ======================================
def extract_payments(erp_df, ven_df):
    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def export_reconciliation_excel(matched, erp_missing, ven_missing, matched_payments, tier2_strict, tier2_relaxed):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not matched.empty: matched.to_excel(writer, sheet_name='Matched_Invoices', index=False)
        if not erp_missing.empty: erp_missing.to_excel(writer, sheet_name='ERP_Missing', index=False)
        if not ven_missing.empty: ven_missing.to_excel(writer, sheet_name='Vendor_Missing', index=False)
        if not tier2_strict.empty: tier2_strict.to_excel(writer, sheet_name='Tier2_Strict', index=False)
        if not tier2_relaxed.empty: tier2_relaxed.to_excel(writer, sheet_name='Tier2_Relaxed', index=False)
        
        summary_data = {
            'Category': ['Perfect Matches', 'Difference Matches', 'Tier-2 Strict', 'Tier-2 Relaxed', 'ERP Unmatched', 'Vendor Unmatched'],
            'Count': [len(matched[matched['Status']=='Perfect Match']), 
                     len(matched[matched['Status']=='Difference Match']), 
                     len(tier2_strict), len(tier2_relaxed), 
                     len(erp_missing), len(ven_missing)]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
    output.seek(0)
    return output.getvalue()

def create_summary_table_with_totals(matched_df, erp_missing, ven_missing):
    erp_total = matched_df['ERP Amount'].sum() + erp_missing['Amount'].sum()
    vendor_total = matched_df['Vendor Amount'].sum() + ven_missing['Amount'].sum()
    matched_erp_total = matched_df['ERP Amount'].sum()
    matched_vendor_total = matched_df['Vendor Amount'].sum()
    total_difference = abs(erp_total - vendor_total)
   
    return pd.DataFrame({
        'Category': ['ERP Total', 'Vendor Total', 'Total Diff', '', 'Matched ERP', 'Matched Vendor', 'Matched Diff', '', 'Unmatched ERP', 'Unmatched Vendor'],
        'Count': [len(matched_df)+len(erp_missing), len(matched_df)+len(ven_missing), '', '', len(matched_df), len(matched_df), '', '', len(erp_missing), len(ven_missing)],
        'Amount': [f"{erp_total:,.2f}", f"{vendor_total:,.2f}", f"{total_difference:,.2f}", '',
                  f"{matched_erp_total:,.2f}", f"{matched_vendor_total:,.2f}", f"{abs(matched_erp_total-matched_vendor_total):,.2f}", '',
                  f"{erp_missing['Amount'].sum():,.2f}", f"{ven_missing['Amount'].sum():,.2f}"]
    })

def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()
  
    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        payment_patterns = [r"^πληρωμ", r"^payment", r"^trf", r"^pago", r"^εξοφληση"]
        if any(re.search(p, reason) for p in payment_patterns): return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό"]
        if any(k in reason for k in credit_words): return "CN"
        return "INV"
  
    def calc_erp_amount(row):
        return abs(normalize_number(row.get("debit_erp", row.get("credit_erp", 0))))
  
    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        payment_patterns = [r"^πληρωμ", r"^payment", r"^trf", r"^pago", r"^εξοφληση"]
        if any(re.search(p, reason) for p in payment_patterns): return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό"]
        if any(k in reason for k in credit_words): return "CN"
        return "INV"
  
    def calc_vendor_amount(row):
        return abs(normalize_number(row.get("debit_ven", row.get("credit_ven", 0))))
  
    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)
  
    erp_use = erp_df[erp_df["__doctype"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__doctype"] != "IGNORE"].copy()
  
    erp_use["__amt"] = erp_use["__amt"].astype(float)
    ven_use["__amt"] = ven_use["__amt"].astype(float)
  
    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            diff = abs(e_amt - v_amt)
            exact_match = (e_inv == v_inv)
            numerical_match = False
            e_nums = re.findall(r'(\d{4,})$', e_inv)
            v_nums = re.findall(r'(\d{4,})$', v_inv)
            if e_nums and v_nums and len(e_nums[0]) == len(v_nums[0]):
                numerical_match = (e_nums[0] == v_nums[0])
            if exact_match or numerical_match:
                if diff <= 0.01:
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
# MAIN UI
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
            
            # TIER-2 STRICT first (exact amounts)
            tier2_strict, used_e, used_v, final_erp_missing, final_ven_missing = tier2_match_strict(erp_missing, ven_missing)
            
            # TIER-2 RELAXED on remaining (NO amount limit)
            tier2_relaxed = tier2_match_relaxed(final_erp_missing, final_ven_missing)
            
            # Final unmatched
            erp_missing = final_erp_missing.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
            ven_missing = final_ven_missing.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
            
        st.success("✅ Reconciliation complete!")
 
        # ======================================
        # EXECUTIVE SUMMARY
        # ======================================
        st.markdown("## 📈 Executive Summary")
        summary_table = create_summary_table_with_totals(matched, erp_missing, ven_missing)
        st.dataframe(summary_table, use_container_width=True, hide_index=True)
 
        # ======================================
        # METRICS
        # ======================================
        col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
        perfect_count = len(matched[matched['Status'] == 'Perfect Match']) if not matched.empty else 0
        diff_count = len(matched[matched['Status'] == 'Difference Match']) if not matched.empty else 0
        
        with col1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("🎯 Perfect", perfect_count)
            st.markdown('</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
            st.metric("⚠️ Differences", diff_count)
            st.markdown('</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="metric-container tier2-strict">', unsafe_allow_html=True)
            st.metric("🔒 Tier-2 Strict", len(tier2_strict))
            st.markdown('</div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="metric-container tier2-relaxed">', unsafe_allow_html=True)
            st.metric("🟡 Tier-2 Relaxed", len(tier2_relaxed))
            st.markdown('</div>', unsafe_allow_html=True)
        with col5:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("❌ ERP Unmatched", len(erp_missing))
            st.markdown('</div>', unsafe_allow_html=True)
        with col6:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("❌ Vendor Unmatched", len(ven_missing))
            st.markdown('</div>', unsafe_allow_html=True)
 
        st.markdown("---")
 
        # ======================================
        # MATCHED INVOICES
        # ======================================
        st.subheader("✅ MATCHED INVOICES")
        if not matched.empty:
            matched_display = matched[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Status']].copy()
            total_row = pd.DataFrame({
                'ERP Invoice': ['TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [matched_display['ERP Amount'].sum()],
                'Vendor Amount': [matched_display['Vendor Amount'].sum()],
                'Difference': [matched_display['Difference'].sum()],
                'Status': [f"TOTAL ({len(matched_display)} MATCHES)"]
            })
            st.dataframe(pd.concat([matched_display, total_row], ignore_index=True), use_container_width=True, height=400)
 
        # ======================================
        # TIER-2 STRICT
        # ======================================
        if not tier2_strict.empty:
            st.markdown("### 🔒 TIER-2 STRICT (Fuzzy Invoice + Exact Amounts ±0.05)")
            tier2s_display = tier2_strict[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Fuzzy Score']].copy()
            total_row_s = pd.DataFrame({
                'ERP Invoice': ['STRICT TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [tier2s_display['ERP Amount'].sum()],
                'Vendor Amount': [tier2s_display['Vendor Amount'].sum()],
                'Difference': [tier2s_display['Difference'].sum()],
                'Fuzzy Score': [f"{len(tier2s_display)} MATCHES"]
            })
            st.dataframe(pd.concat([tier2s_display, total_row_s], ignore_index=True), use_container_width=True)
 
        # ======================================
        # TIER-2 RELAXED
        # ======================================
        if not tier2_relaxed.empty:
            st.markdown("### 🟡 TIER-2 RELAXED (Fuzzy Invoice Pattern - ANY Amount Difference)")
            st.info("💡 **Invoice numbers have similar patterns** - Review amount differences manually")
            
            tier2r_display = tier2_relaxed[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Amount Diff', 'Fuzzy Score']].copy()
            total_row_r = pd.DataFrame({
                'ERP Invoice': ['RELAXED TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [tier2r_display['ERP Amount'].sum()],
                'Vendor Amount': [tier2r_display['Vendor Amount'].sum()],
                'Amount Diff': [tier2r_display['Amount Diff'].sum()],
                'Fuzzy Score': [f"{len(tier2r_display)} MATCHES"]
            })
            st.dataframe(pd.concat([tier2r_display, total_row_r], ignore_index=True), use_container_width=True, height=400)
            st.warning(f"**{len(tier2_relaxed)} RELAXED MATCHES | Total Amount Diff: €{tier2_relaxed['Amount Diff'].sum():,.2f}**")
 
        # ======================================
        # UNMATCHED
        # ======================================
        st.subheader("❌ UNMATCHED INVOICES")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**🔴 Missing in ERP**")
            if not ven_missing.empty:
                st.dataframe(ven_missing, use_container_width=True)
                st.error(f"**{len(ven_missing)} UNMATCHED | €{ven_missing['Amount'].sum():,.2f}**")
        with col2:
            st.markdown("**🔴 Missing in Vendor**")
            if not erp_missing.empty:
                st.dataframe(erp_missing, use_container_width=True)
                st.error(f"**{len(erp_missing)} UNMATCHED | €{erp_missing['Amount'].sum():,.2f}**")
 
        # ======================================
        # DOWNLOAD
        # ======================================
        st.markdown("### 📥 Download Full Report")
        excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing, pd.DataFrame(), tier2_strict, tier2_relaxed)
        st.download_button(
            "💾 Download Excel Report",
            data=excel_output,
            file_name="ReconRaptor_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
       
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Check your Excel files have invoice/amount/date columns")
