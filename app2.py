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
    .tier2-amount-diff { background-color: #FFCA28 !important; color: black !important; font-weight: bold !important; }
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
@st.cache_data
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
    try:
        return float(s)
    except:
        return 0.0

@st.cache_data
def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "": return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if not pd.isna(d): return d.strftime("%Y-%m-%d")
    except: pass
    return ""

@st.cache_data
def clean_invoice_code(v):
    if not v: return "0"
    s = str(v).strip().lower()
    # Extract the MAIN numeric sequence (longest one)
    digits = re.findall(r'\b\d{4,}\b', s)  # Word boundary + 4+ digits
    if digits:
        return digits[-1].lstrip('0') or '0'
    # Fallback to any digits
    digits = re.findall(r'\d{3,}', s)
    if digits:
        return digits[-1].lstrip('0') or '0'
    return "0"

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
# TIER-2 STRICT (UNCHANGED - FAST)
# ======================================
def tier2_match_strict(erp_missing, ven_missing):
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), set(), set(), erp_missing.copy(), ven_missing.copy()
    
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    
    e_codes = e_df["invoice_erp"].apply(clean_invoice_code).values
    v_codes = v_df["invoice_ven"].apply(clean_invoice_code).values
    e_amts = e_df["__amt"].astype(float).values
    v_amts = v_df["__amt"].astype(float).values
    
    matches = []
    used_e, used_v = set(), set()
    
    for i, e_code in enumerate(e_codes):
        if i in used_e: continue
        e_amt = e_amts[i]
        
        # Only same length codes
        candidates = [j for j, v_code in enumerate(v_codes) if len(str(e_code)) == len(str(v_code)) and j not in used_v]
        
        for j in candidates[:5]:  # Top 5 candidates
            v_amt = v_amts[j]
            if abs(e_amt - v_amt) > 0.05: continue
            
            sim = SequenceMatcher(None, str(e_code), str(v_codes[j])).ratio()
            if sim >= 0.8:
                matches.append({
                    "ERP Invoice": e_df.iloc[i]["invoice_erp"],
                    "Vendor Invoice": v_df.iloc[j]["invoice_ven"],
                    "ERP Amount": round(e_amt, 2),
                    "Vendor Amount": round(v_amt, 2),
                    "Difference": abs(e_amt - v_amt),
                    "Fuzzy Score": f"{sim:.1%}",
                    "Match Type": "🔒 Tier-2 Strict"
                })
                used_e.add(i)
                used_v.add(j)
                break
    
    tier2_matches = pd.DataFrame(matches)
    remaining_erp = e_df.drop(list(used_e)).reset_index(drop=True)
    remaining_ven = v_df.drop(list(used_v)).reset_index(drop=True)
    
    return tier2_matches, used_e, used_v, remaining_erp, remaining_ven

# ======================================
# NEW: TIER-2 AMOUNT DIFFERENCE (90%+ FUZZY INVOICE)
# ======================================
def tier2_match_amount_diff(erp_missing, ven_missing):
    """🟡 TIER-2 AMOUNT DIFF: 90%+ FUZZY INVOICE + ANY amount difference"""
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame()
    
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    
    e_codes = e_df["invoice_erp"].apply(clean_invoice_code).values
    v_codes = v_df["invoice_ven"].apply(clean_invoice_code).values
    e_amts = e_df["__amt"].astype(float).values
    v_amts = v_df["__amt"].astype(float).values
    
    matches = []
    
    # BUCKET BY LENGTH (SUPER FAST)
    code_len_to_indices = {}
    for j, v_code in enumerate(v_codes):
        length = len(str(v_code))
        if length not in code_len_to_indices:
            code_len_to_indices[length] = []
        code_len_to_indices[length].append(j)
    
    for i, e_code in enumerate(e_codes):
        e_code_str = str(e_code)
        length = len(e_code_str)
        
        if length not in code_len_to_indices:
            continue
            
        # Top 3 candidates per bucket
        candidates = code_len_to_indices[length][:3]
        
        for j in candidates:
            v_code_str = str(v_codes[j])
            sim = SequenceMatcher(None, e_code_str, v_code_str).ratio()
            
            # **90%+ FUZZY MATCH** - GOOD PATTERN SIMILARITY
            if sim >= 0.90:
                matches.append({
                    "ERP Invoice": e_df.iloc[i]["invoice_erp"],
                    "Vendor Invoice": v_df.iloc[j]["invoice_ven"],
                    "ERP Amount": round(e_amts[i], 2),
                    "Vendor Amount": round(v_amts[j], 2),
                    "Amount Diff": abs(e_amts[i] - v_amts[j]),
                    "Fuzzy Score": f"{sim:.1%}",
                    "ERP Date": e_df.iloc[i].get("date_erp", ""),
                    "Vendor Date": v_df.iloc[j].get("date_ven", ""),
                    "Match Type": "🟡 Tier-2 Amount Diff"
                })
                break  # One match per ERP invoice
    
    return pd.DataFrame(matches)

# ======================================
# MATCH_INVOICES (FAST VERSION)
# ======================================
@st.cache_data
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
# UTILITY FUNCTIONS
# ======================================
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

def export_reconciliation_excel(matched, erp_missing, ven_missing, tier2_strict, tier2_amount_diff):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not matched.empty: matched.to_excel(writer, sheet_name='Matched_Invoices', index=False)
        if not erp_missing.empty: erp_missing.to_excel(writer, sheet_name='ERP_Missing', index=False)
        if not ven_missing.empty: ven_missing.to_excel(writer, sheet_name='Vendor_Missing', index=False)
        if not tier2_strict.empty: tier2_strict.to_excel(writer, sheet_name='Tier2_Strict', index=False)
        if not tier2_amount_diff.empty: tier2_amount_diff.to_excel(writer, sheet_name='Tier2_Amount_Diff', index=False)
       
        summary_data = {
            'Category': ['Perfect Matches', 'Difference Matches', 'Tier-2 Strict', 'Tier-2 Amount Diff', 'ERP Unmatched', 'Vendor Unmatched'],
            'Count': [len(matched[matched['Status']=='Perfect Match']) if not matched.empty else 0,
                     len(matched[matched['Status']=='Difference Match']) if not matched.empty else 0,
                     len(tier2_strict), len(tier2_amount_diff),
                     len(erp_missing), len(ven_missing)]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
    output.seek(0)
    return output.getvalue()

# ======================================
# MAIN UI
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        with st.spinner("🔍 Analyzing invoices..."):
            erp_raw = pd.read_excel(uploaded_erp, dtype=str)
            ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
            erp_df = normalize_columns(erp_raw, "erp")
            ven_df = normalize_columns(ven_raw, "ven")
            
            # Step 1: Exact matches
            matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
            
            # Step 2: Tier-2 Strict (exact amounts)
            tier2_strict, used_e, used_v, final_erp_missing, final_ven_missing = tier2_match_strict(erp_missing, ven_missing)
            
            # Step 3: Tier-2 Amount Diff (90%+ fuzzy invoice)
            tier2_amount_diff = tier2_match_amount_diff(final_erp_missing, final_ven_missing)
            
            # Final unmatched
            erp_missing = final_erp_missing.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
            ven_missing = final_ven_missing.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
        
        st.success("✅ Analysis complete!")
        
        # ======================================
        # EXECUTIVE SUMMARY
        # ======================================
        st.markdown("## 📈 Executive Summary")
        summary_table = create_summary_table_with_totals(matched, erp_missing, ven_missing)
        st.dataframe(summary_table, use_container_width=True, hide_index=True)
        
        # ======================================
        # METRICS
        # ======================================
        col1, col2, col3, col4, col5, col6 = st.columns(6)
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
            st.markdown('<div class="metric-container tier2-amount-diff">', unsafe_allow_html=True)
            st.metric("🟡 Tier-2 Amount Diff", len(tier2_amount_diff))
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
            st.markdown("### 🔒 TIER-2 STRICT")
            st.info("💡 Fuzzy invoice numbers + **EXACT amounts** (±0.05)")
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
        # TIER-2 AMOUNT DIFFERENCE (YOUR REQUESTED TABLE)
        # ======================================
        if not tier2_amount_diff.empty:
            st.markdown("### 🟡 TIER-2 AMOUNT DIFFERENCE")
            st.info("💡 **90%+ FUZZY INVOICE MATCH** + **DIFFERENT amounts** - Review these!")
            
            tier2ad_display = tier2_amount_diff[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Amount Diff', 'Fuzzy Score']].copy()
            total_row_ad = pd.DataFrame({
                'ERP Invoice': ['AMOUNT DIFF TOTAL'],
                'Vendor Invoice': [''],
                'ERP Amount': [tier2ad_display['ERP Amount'].sum()],
                'Vendor Amount': [tier2ad_display['Vendor Amount'].sum()],
                'Amount Diff': [tier2ad_display['Amount Diff'].sum()],
                'Fuzzy Score': [f"{len(tier2ad_display)} MATCHES"]
            })
            st.dataframe(pd.concat([tier2ad_display, total_row_ad], ignore_index=True), use_container_width=True, height=400)
            st.warning(f"**{len(tier2_amount_diff)} FUZZY INVOICES | Total Amount Diff: €{tier2_amount_diff['Amount Diff'].sum():,.2f}**")
        
        # ======================================
        # UNMATCHED
        # ======================================
        st.subheader("❌ UNMATCHED INVOICES")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**🔴 Missing in ERP (Vendor Only)**")
            if not ven_missing.empty:
                st.dataframe(ven_missing, use_container_width=True)
                st.error(f"**{len(ven_missing)} UNMATCHED | €{ven_missing['Amount'].sum():,.2f}**")
            else:
                st.success("✅ No unmatched vendor invoices!")
        with col2:
            st.markdown("**🔴 Missing in Vendor (ERP Only)**")
            if not erp_missing.empty:
                st.dataframe(erp_missing, use_container_width=True)
                st.error(f"**{len(erp_missing)} UNMATCHED | €{erp_missing['Amount'].sum():,.2f}**")
            else:
                st.success("✅ No unmatched ERP invoices!")
        
        # ======================================
        # DOWNLOAD
        # ======================================
        st.markdown("### 📥 Download Full Report")
        excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing, tier2_strict, tier2_amount_diff)
        st.download_button(
            "💾 Download Excel Report",
            data=excel_output,
            file_name="ReconRaptor_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
      
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Check your Excel files have invoice/amount/date columns")
