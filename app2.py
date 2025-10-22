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
    .missing-erp { background-color: #C62828 !important; color: white !important; font-weight: bold !important; }
    .missing-vendor { background-color: #AD1457 !important; color: white !important; font-weight: bold !important; }
    .payment-match { background-color: #004D40 !important; color: white !important; font-weight: bold !important; }
    .erp-payment { background-color: #4CAF50 !important; color: white !important; }
    .vendor-payment { background-color: #2196F3 !important; color: white !important; }
    .metric-container { padding: 1rem !important; border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)

st.title("🦖 ReconRaptor — Vendor Reconciliation")

# ======================================
# HELPERS (unchanged)
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
        "%m/%d/%y", "%m-%d-%y",
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
# ENHANCED COLOR STYLING FUNCTIONS
# ======================================
@st.cache_data
def colorize_dataframe(df, status_col=None):
    """Apply color styling to dataframe based on status"""
    def apply_color(val, row=None):
        if status_col and row is not None:
            status = row.get(status_col, '')
            if status == 'Perfect Match':
                return f'<span class="perfect-match">{val}</span>'
            elif status == 'Difference Match':
                return f'<span class="difference-match">{val}</span>'
            elif status == 'Tier-2':
                return f'<span class="tier2-match">{val}</span>'
        return val
    
    if not df.empty and status_col in df.columns:
        styled_df = df.copy()
        html_df = df.style.apply(lambda row: [apply_color(val, row) for val in row], axis=1).to_html()
        return html_df
    return df.to_html()

def style_perfect_matches(df):
    """Style perfect matches - GREEN"""
    return df.style.apply(
        lambda row: ['background-color: #2E7D32; color: white; font-weight: bold'] * len(row), 
        axis=1
    )

def style_difference_matches(df):
    """Style difference matches - YELLOW"""
    return df.style.apply(
        lambda row: ['background-color: #F9A825; color: black; font-weight: bold'] * len(row), 
        axis=1
    )

def style_tier2_matches(df):
    """Style tier-2 matches - TEAL"""
    return df.style.apply(
        lambda row: ['background-color: #26A69A; color: white; font-weight: bold'] * len(row), 
        axis=1
    )

def style_missing(df):
    """Style missing - RED"""
    return df.style.apply(
        lambda row: ['background-color: #C62828; color: white; font-weight: bold'] * len(row), 
        axis=1
    )

# [Keep all other functions unchanged - match_invoices, tier2_match, extract_payments, export_reconciliation_excel]
# ... (I'll include them in the full code at the end)

# ======================================
# STREAMLIT UI WITH FULL COLORS
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")
  
    with st.spinner("🔍 Analyzing and reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)
        tier2_matches, used_erp_indices, used_ven_indices, _, _ = tier2_match(erp_missing, ven_missing)
      
        if used_erp_indices:
            erp_missing = erp_missing[~erp_missing.index.isin(used_erp_indices)]
        if used_ven_indices:
            ven_missing = ven_missing[~ven_missing.index.isin(used_ven_indices)]
  
    st.success("✅ Reconciliation complete!")
  
    # SUMMARY METRICS WITH COLORS
    col1, col2, col3, col4 = st.columns(4)
    perfect_count = len(matched[matched['Status'] == 'Perfect Match']) if not matched.empty else 0
    diff_count = len(matched[matched['Status'] == 'Difference Match']) if not matched.empty else 0
    tier2_count = len(tier2_matches) if not tier2_matches.empty else 0
   
    with col1:
        st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
        st.metric("🎯 Perfect Matches", perfect_count, delta=None)
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
        st.metric("⚠️ Differences", diff_count, delta=None)
        st.markdown('</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
        st.metric("🔍 Tier-2 Matches", tier2_count, delta=None)
        st.markdown('</div>', unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="metric-container payment-match">', unsafe_allow_html=True)
        st.metric("✅ Total Reconciled", perfect_count + diff_count + tier2_count, delta=None)
        st.markdown('</div>', unsafe_allow_html=True)
  
    st.markdown("---")
  
    # TIER-1 RESULTS WITH COLORS
    st.subheader("📊 Tier-1 Matches & Differences")
    if not matched.empty:
        perfect_matches = matched[matched['Status'] == 'Perfect Match']
        diff_matches = matched[matched['Status'] == 'Difference Match']
       
        col1, col2 = st.columns(2)
       
        with col1:
            st.markdown("**✅ Perfect Matches** 🟢")
            if not perfect_matches.empty:
                st.dataframe(
                    style_perfect_matches(perfect_matches[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Status']]),
                    use_container_width=True,
                    height=400
                )
            else:
                st.info("No perfect matches found.")
       
        with col2:
            st.markdown("**⚠️ Amount Differences** 🟡")
            if not diff_matches.empty:
                st.dataframe(
                    style_difference_matches(diff_matches[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Status']]),
                    use_container_width=True,
                    height=400
                )
            else:
                st.success("No amount differences found!")
    else:
        st.info("❌ No Tier-1 matches/differences found.")
  
    # TIER-2 WITH COLORS
    st.subheader("🔍 Tier-2 Matches (Fuzzy)")
    if not tier2_matches.empty:
        styled_tier2 = tier2_matches.copy()
        styled_tier2['Match Type'] = 'Tier-2'
        st.dataframe(
            style_tier2_matches(styled_tier2[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference', 'Fuzzy Score', 'Match Type']]),
            use_container_width=True
        )
    else:
        st.info("No Tier-2 matches found.")
  
    # MISSING WITH COLORS
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### ❌ Missing in ERP 🔴")
        if not erp_missing.empty:
            styled_missing_erp = erp_missing.copy()
            styled_missing_erp['Status'] = 'Missing ERP'
            st.dataframe(
                style_missing(styled_missing_erp),
                use_container_width=True
            )
            st.error(f"**{len(erp_missing)} invoices** missing in ERP")
        else:
            st.success("✅ No missing invoices in ERP")
   
    with col2:
        st.markdown("### ❌ Missing in Vendor 🔴")
        if not ven_missing.empty:
            styled_missing_ven = ven_missing.copy()
            styled_missing_ven['Status'] = 'Missing Vendor'
            st.dataframe(
                style_missing(styled_missing_ven),
                use_container_width=True
            )
            st.error(f"**{len(ven_missing)} invoices** missing in Vendor")
        else:
            st.success("✅ No missing invoices in Vendor")
  
    # PAYMENTS WITH COLORS
    st.subheader("🏦 Payment Transactions")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**💼 ERP Payments** 🟢")
        if not erp_pay.empty:
            display_erp = erp_pay[['reason_erp', 'debit_erp', 'credit_erp', 'Amount']].copy()
            display_erp.columns = ['Reason', 'Debit', 'Credit', 'Net Amount']
            st.dataframe(
                display_erp.style.apply(lambda row: ['background-color: #4CAF50; color: white'] * len(row), axis=1),
                use_container_width=True
            )
            st.markdown(f"**Total: {erp_pay['Amount'].sum():,.2f} EUR**")
        else:
            st.info("No ERP payments found.")
  
    with col2:
        st.markdown("**🧾 Vendor Payments** 🔵")
        if not ven_pay.empty:
            display_ven = ven_pay[['reason_ven', 'debit_ven', 'credit_ven', 'Amount']].copy()
            display_ven.columns = ['Reason', 'Debit', 'Credit', 'Net Amount']
            st.dataframe(
                display_ven.style.apply(lambda row: ['background-color: #2196F3; color: white'] * len(row), axis=1),
                use_container_width=True
            )
            st.markdown(f"**Total: {ven_pay['Amount'].sum():,.2f} EUR**")
        else:
            st.info("No Vendor payments found.")
  
    if not matched_pay.empty:
        st.subheader("✅ Matched Payments 🟢")
        st.dataframe(
            matched_pay.style.apply(lambda row: ['background-color: #004D40; color: white; font-weight: bold'] * len(row), axis=1),
            use_container_width=True
        )
  
    # Download
    st.markdown("### 📥 Download Full Report")
    excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing, matched_pay, tier2_matches)
    st.download_button(
        "💾 Download Excel Report",
        data=excel_output,
        file_name="ReconRaptor_Reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
