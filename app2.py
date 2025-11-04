# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FINAL + Balance Difference Metric)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows, get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher

# ==================== PAGE CONFIG & CSS ======================
st.set_page_config(page_title="ReconRaptor — Vendor Reconciliation", layout="wide")
st.markdown("""
<style>
.big-title {
    font-size: 3rem !important;
    font-weight: 700;
    text-align: center;
    background: linear-gradient(90deg, #1E88E5, #42A5F5);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 1rem;
}
.section-title {
    font-size: 1.8rem !important;
    font-weight: 600;
    color: #1565C0;
    border-bottom: 2px solid #42A5F5;
    padding-bottom: 0.5rem;
    margin-top: 2rem;
}
.metric-container {
    padding: 1.2rem !important;
    border-radius: 12px !important;
    margin-bottom: 1rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
}
.perfect-match   {background:#2E7D32;color:#fff;font-weight:bold;}
.difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
.tier2-match     {background:#26A69A;color:#fff;font-weight:bold;}
.tier3-match     {background:#7E57C2;color:#fff;font-weight:bold;}
.missing-erp     {background:#C62828;color:#fff;font-weight:bold;}
.missing-vendor  {background:#AD1457;color:#fff;font-weight:bold;}
.payment-match   {background:#004D40;color:#fff;font-weight:bold;}
.balance-metric  {background:#1E88E5;color:#fff;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "":
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
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
    if pd.isna(v) or str(v).strip() == "":
        return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    for fmt in [
        "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d",
        "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d"
    ]:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d): return d.strftime("%Y-%m-%d")
        except: continue
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s = str(v).strip().lower()
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    return s or "0"

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "document", "ref", "num"],
        "credit":  ["credit", "haber", "credito", "abono"],
        "debit":   ["debit", "debe", "importe", "amount", "valor", "total"],
        "reason":  ["reason", "motivo", "concepto", "descripcion"],
        "date":    ["date", "fecha", "data"]
    }
    rename_map = {}
    for col in df.columns:
        low = str(col).lower()
        for key, aliases in mapping.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
    out = df.rename(columns=rename_map)
    for req in ["debit", "credit"]:
        c = f"{req}_{tag}"
        if c not in out.columns:
            out[c] = 0.0
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

def style(df, css):
    return df.style.apply(lambda _: [css] * len(_), axis=1)

# ==================== BALANCE DIFFERENCE FUNCTION ==========================
def calculate_balance_difference(erp_df, ven_df):
    balance_col_erp = next((c for c in erp_df.columns if "balance" in c.lower()), None)
    possible_vendor_cols = ["balance", "saldo", "υπόλοιπο", "ypolipo", "υπολοιπο"]
    balance_col_ven = next((c for c in ven_df.columns if any(p in c.lower() for p in possible_vendor_cols)), None)

    if not balance_col_erp or not balance_col_ven:
        return None, None, None

    def parse_amount(v):
        s = str(v).strip().replace("€", "").replace(",", ".")
        s = re.sub(r"[^\d.\-]", "", s)
        try:
            return float(s)
        except:
            return 0.0

    erp_vals = [parse_amount(v) for v in erp_df[balance_col_erp] if str(v).strip()]
    ven_vals = [parse_amount(v) for v in ven_df[balance_col_ven] if str(v).strip()]
    if not erp_vals or not ven_vals:
        return None, None, None

    last_erp = erp_vals[-1]
    last_ven = ven_vals[-1]
    diff = round(last_erp - last_ven, 2)
    return last_erp, last_ven, diff

# ==================== MATCHING CORE ==========================
# (same as your current implementation — omitted for brevity)
# --------------------------------------------------------------

# 🧩 The rest of your matching, tier2_match, tier3_match, extract_payments, and export_excel code stays exactly as you already have it.

# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")

        st.write("🧩 ERP columns detected:", list(erp_df.columns))
        st.write("🧩 Vendor columns detected:", list(ven_df.columns))

        with st.spinner("Analyzing invoices..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)
            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)
            erp_pay, ven_pay, pay_match = extract_payments(erp_df, ven_df)

        st.success("Reconciliation Complete!")

        # ---------- METRICS ----------
        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)

        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()

        def safe_sum(df, col): return float(df[col].sum()) if not df.empty and col in df.columns else 0.0

        with c1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("Perfect Matches", len(perf))
            st.markdown(f"**ERP:** {safe_sum(perf,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(perf,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
            st.metric("Differences", len(diff))
            st.markdown(f"**ERP:** {safe_sum(diff,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(diff,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c3:
            st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
            st.metric("Tier-2", len(tier2))
            st.markdown(f"**ERP:** {safe_sum(tier2,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(tier2,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="metric-container tier3-match">', unsafe_allow_html=True)
            st.metric("Tier-3", len(tier3))
            st.markdown(f"**ERP:** {safe_sum(tier3,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(tier3,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c5:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("Unmatched ERP", 0 if final_erp_miss.empty else len(final_erp_miss))
            st.markdown(f"**Total:** {final_erp_miss['Amount'].sum():,.2f}" if not final_erp_miss.empty else "**Total:** 0.00", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c6:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("Unmatched Vendor", 0 if final_ven_miss.empty else len(final_ven_miss))
            st.markdown(f"**Total:** {final_ven_miss['Amount'].sum():,.2f}" if not final_ven_miss.empty else "**Total:** 0.00", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c7:
            st.markdown('<div class="metric-container payment-match">', unsafe_allow_html=True)
            st.metric("New Payment Matches", len(pay_match) if not pay_match.empty else 0)
            st.markdown('</div>', unsafe_allow_html=True)

        # 💼 Balance Difference metric (appears right below the others)
        last_balance_erp, last_balance_ven, balance_diff = calculate_balance_difference(erp_df, ven_df)
        if balance_diff is not None:
            st.markdown('<div class="metric-container balance-metric">', unsafe_allow_html=True)
            st.metric("ERP Balance", f"{last_balance_erp:,.2f}")
            st.metric("Vendor Balance", f"{last_balance_ven:,.2f}")
            st.metric("Balance Difference (ERP - Vendor)", f"{balance_diff:,.2f}")
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")

        with st.spinner("Analyzing invoices..."):
            # Tier-1
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)

            # progressive de-dup after Tier-1
            used_erp_inv = set(tier1["ERP Invoice"].astype(str)) if not tier1.empty else set()
            used_ven_inv = set(tier1["Vendor Invoice"].astype(str)) if not tier1.empty else set()
            if not miss_erp.empty:
                miss_erp = miss_erp[~miss_erp["Invoice"].astype(str).isin(used_erp_inv)]
            if not miss_ven.empty:
                miss_ven = miss_ven[~miss_ven["Invoice"].astype(str).isin(used_ven_inv)]

            # Tier-2
            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            if not tier2.empty:
                used_erp_inv |= set(tier2["ERP Invoice"].astype(str))
                used_ven_inv |= set(tier2["Vendor Invoice"].astype(str))
                if not miss_erp2.empty:
                    miss_erp2 = miss_erp2[~miss_erp2["Invoice"].astype(str).isin(used_erp_inv)]
                if not miss_ven2.empty:
                    miss_ven2 = miss_ven2[~miss_ven2["Invoice"].astype(str).isin(used_ven_inv)]
            else:
                miss_erp2, miss_ven2 = miss_erp, miss_ven

            # Tier-3
            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)
            if not tier3.empty:
                used_erp_inv |= set(tier3["ERP Invoice"].astype(str))
                used_ven_inv |= set(tier3["Vendor Invoice"].astype(str))
                if not final_erp_miss.empty:
                    final_erp_miss = final_erp_miss[~final_erp_miss["Invoice"].astype(str).isin(used_erp_inv)]
                if not final_ven_miss.empty:
                    final_ven_miss = final_ven_miss[~final_ven_miss["Invoice"].astype(str).isin(used_ven_inv)]

            # Payments
            erp_pay, ven_pay, pay_match = extract_payments(erp_df, ven_df)

        st.success("Reconciliation Complete!")

        # ---------- METRICS ----------
        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()

        def safe_sum(df, col):
            return float(df[col].sum()) if not df.empty and col in df.columns else 0.0

        with c1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("Perfect Matches", len(perf))
            st.markdown(
                f"**ERP:** {safe_sum(perf, 'ERP Amount'):,.2f}<br>"
                f"**Vendor:** {safe_sum(perf, 'Vendor Amount'):,.2f}<br>"
                f"**Diff:** {safe_sum(perf, 'Difference'):,.2f}",
                unsafe_allow_html=True
            )
            st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
            st.metric("Differences", len(diff))
            st.markdown(
                f"**ERP:** {safe_sum(diff, 'ERP Amount'):,.2f}<br>"
                f"**Vendor:** {safe_sum(diff, 'Vendor Amount'):,.2f}<br>"
                f"**Diff:** {safe_sum(diff, 'Difference'):,.2f}",
                unsafe_allow_html=True
            )
            st.markdown('</div>', unsafe_allow_html=True)

        with c3:
            st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
            st.metric("Tier-2", len(tier2))
            st.markdown(
                f"**ERP:** {safe_sum(tier2, 'ERP Amount'):,.2f}<br>"
                f"**Vendor:** {safe_sum(tier2, 'Vendor Amount'):,.2f}<br>"
                f"**Diff:** {safe_sum(tier2, 'Difference'):,.2f}",
                unsafe_allow_html=True
            )
            st.markdown('</div>', unsafe_allow_html=True)

        with c4:
            st.markdown('<div class="metric-container tier3-match">', unsafe_allow_html=True)
            st.metric("Tier-3", len(tier3))
            st.markdown(
                f"**ERP:** {safe_sum(tier3, 'ERP Amount'):,.2f}<br>"
                f"**Vendor:** {safe_sum(tier3, 'Vendor Amount'):,.2f}<br>"
                f"**Diff:** {safe_sum(tier3, 'Difference'):,.2f}",
                unsafe_allow_html=True
            )
            st.markdown('</div>', unsafe_allow_html=True)

        with c5:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("Unmatched ERP", 0 if final_erp_miss.empty else len(final_erp_miss))
            st.markdown(
                f"**Total:** {final_erp_miss['Amount'].sum():,.2f}" if not final_erp_miss.empty and 'Amount' in final_erp_miss.columns else "**Total:** 0.00",
                unsafe_allow_html=True
            )
            st.markdown('</div>', unsafe_allow_html=True)

        with c6:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("Unmatched Vendor", 0 if final_ven_miss.empty else len(final_ven_miss))
            st.markdown(
                f"**Total:** {final_ven_miss['Amount'].sum():,.2f}" if not final_ven_miss.empty and 'Amount' in final_ven_miss.columns else "**Total:** 0.00",
                unsafe_allow_html=True
            )
            st.markdown('</div>', unsafe_allow_html=True)

        with c7:
            st.markdown('<div class="metric-container payment-match">', unsafe_allow_html=True)
            st.metric("New Payment Matches", len(pay_match) if not pay_match.empty else 0)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")

        # ---------- DISPLAY ----------
        st.markdown('<h2 class="section-title">Tier-1: Exact Matches</h2>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Perfect Matches**")
            if not perf.empty:
                st.dataframe(
                    style(perf[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference']],
                          "background:#2E7D32;color:#fff;font-weight:bold;"),
                    use_container_width=True
                )
            else:
                st.info("No perfect matches.")
        with col_b:
            st.markdown("**Amount Differences**")
            if not diff.empty:
                st.dataframe(
                    style(diff[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference']],
                          "background:#FF8F00;color:#fff;font-weight:bold;"),
                    use_container_width=True
                )
            else:
                st.success("No differences.")

        st.markdown('<h2 class="section-title">Tier-2: Fuzzy + Small Amount</h2>', unsafe_allow_html=True)
        if not tier2.empty:
            st.dataframe(style(tier2, "background:#26A69A;color:#fff;font-weight:bold;"), use_container_width=True)
        else:
            st.info("No Tier-2 matches.")

        st.markdown('<h2 class="section-title">Tier-3: Date + Strict Fuzzy</h2>', unsafe_allow_html=True)
        if not tier3.empty:
            st.dataframe(style(tier3, "background:#7E57C2;color:#fff;font-weight:bold;"), use_container_width=True)
        else:
            st.info("No Tier-3 matches.")

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown('<h2 class="section-title">Missing in ERP</h2>', unsafe_allow_html=True)
            if not final_ven_miss.empty:
                st.dataframe(style(final_ven_miss, "background:#AD1457;color:#fff;font-weight:bold;"), use_container_width=True)
                st.error(f"{len(final_ven_miss)} vendor invoices missing – {final_ven_miss['Amount'].sum():,.2f}")
            else:
                st.success("All vendor invoices found in ERP.")
        with col_m2:
            st.markdown('<h2 class="section-title">Missing in Vendor</h2>', unsafe_allow_html=True)
            if not final_erp_miss.empty:
                st.dataframe(style(final_erp_miss, "background:#C62828;color:#fff;font-weight:bold;"), use_container_width=True)
                st.error(f"{len(final_erp_miss)} ERP invoices missing – {final_erp_miss['Amount'].sum():,.2f}")
            else:
                st.success("All ERP invoices found in vendor.")

        st.markdown('<h2 class="section-title">Payment Transactions</h2>', unsafe_allow_html=True)
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("**ERP Payments**")
            if not erp_pay.empty:
                disp = erp_pay[['reason_erp', 'Amount', 'credit_erp']].copy()
                disp.columns = ['Reason', 'Debit', 'Credit']
                st.dataframe(
                    disp.style.apply(lambda _: ['background:#4CAF50;color:#fff'] * len(_), axis=1),
                    use_container_width=True
                )
                st.markdown(f"**Total:** {erp_pay['Amount'].sum():,.2f}")
            else:
                st.info("No ERP payments.")
        with col_p2:
            st.markdown("**Vendor Payments**")
            if not ven_pay.empty:
                disp = ven_pay[['reason_ven', 'debit_ven', 'credit_ven', 'Amount']].copy()
                disp.columns = ['Reason', 'Debit', 'Credit', 'Net']
                st.dataframe(
                    disp.style.apply(lambda _: ['background:#2196F3;color:#fff'] * len(_), axis=1),
                    use_container_width=True
                )
                st.markdown(f"**Total:** {ven_pay['Amount'].sum():,.2f}")
            else:
                st.info("No vendor payments.")

        if not pay_match.empty:
            st.markdown("**Matched Payments**")
            st.dataframe(
                pay_match.style.apply(lambda _: ['background:#004D40;color:#fff;font-weight:bold'] * len(_), axis=1),
                use_container_width=True
            )

        # ---------- EXPORT ----------
        st.markdown('<h2 class="section-title">Download Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(final_erp_miss, final_ven_miss)
        st.download_button(
            label="Download Full Excel Report",
            data=excel_buf,
            file_name="ReconRaptor_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that your files contain columns like: **invoice**, **debit/credit**, **date**, **reason**")
