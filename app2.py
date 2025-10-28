# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (Enhanced Build)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
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
.metric-box {
    border-radius: 12px;
    padding: 1.5rem;
    margin: 0.5rem;
    text-align: center;
    color: white;
    font-weight: 600;
    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
}
.green {background: #2E7D32;}
.orange {background: #FF8F00;}
.teal {background: #26A69A;}
.purple {background: #7E57C2;}
.red {background: #C62828;}
.pink {background: #AD1457;}
.dark {background: #004D40;}
</style>
""", unsafe_allow_html=True)
st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)
# ====================== HELPERS ==========================
def fuzzy_ratio(a, b): return SequenceMatcher(None, str(a), str(b)).ratio()
def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "":
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if s.count(",") == 1 and s.count(".") == 1:
        s = s.replace(".", "").replace(",", ".") if s.find(",") > s.find(".") else s.replace(",", "")
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
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(d):
        d = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""
def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "número", "document",
                    "doc", "ref", "referencia", "nº factura", "num factura"],
        "credit": ["credit", "haber", "credito", "crédito", "abono"],
        "debit": ["debit", "debe", "cargo", "importe", "valor", "amount", "total"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "detalle", "descripción"],
        "date": ["date", "fecha", "fech", "data", "issue date", "posting date"]
    }
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
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
# ==================== CONSOLIDATION ==========================
def consolidate_by_invoice(df: pd.DataFrame, inv_col: str) -> pd.DataFrame:
    if inv_col not in df.columns:
        return pd.DataFrame(columns=df.columns)
    tag = "erp" if "erp" in inv_col else "ven"
    debit_col = f"debit_{tag}"
    credit_col = f"credit_{tag}"
    date_col = f"date_{tag}"
    df['__id'] = df.index
    with_inv = df[df[inv_col].notna() & (df[inv_col] != '')]
    without_inv = df[df[inv_col].isna() | (df[inv_col] == '')]
    records = []
    for inv, group in with_inv.groupby(inv_col):
        total_debit = group[debit_col].apply(normalize_number).sum()
        total_credit = group[credit_col].apply(normalize_number).sum()
        net = round(total_debit - total_credit, 2)
        base = group.iloc[0].copy()
        base["__amt"] = abs(net)
        base["__type"] = "INV" if net >= 0 else "CN"
        base[debit_col] = max(net, 0)
        base[credit_col] = -min(net, 0)
        base['__group_ids'] = list(group['__id'])
        records.append(base)
    for idx, row in without_inv.iterrows():
        debit = normalize_number(row[debit_col])
        credit = normalize_number(row[credit_col])
        net = round(debit - credit, 2)
        row["__amt"] = abs(net)
        row["__type"] = "PAY" if net < 0 else "OTHER"
        row[debit_col] = max(net, 0)
        row[credit_col] = -min(net, 0)
        row['__group_ids'] = [row['__id']]
        records.append(row)
    cons_df = pd.DataFrame(records)
    cons_df = cons_df[cons_df['__type'].isin(['INV', 'CN', 'PAY'])]
    return cons_df.reset_index(drop=True)
# ==================== MATCHING CORE ==========================
def match_invoices(erp_use, ven_use):
    inv_erp = erp_use[erp_use['__type'].isin(['INV','CN'])]
    pay_erp = erp_use[erp_use['__type']=='PAY']
    inv_ven = ven_use[ven_use['__type'].isin(['INV','CN'])]
    pay_ven = ven_use[ven_use['__type']=='PAY']
    matched_perfect = []
    matched_diff = []
    matched_tier3 = []
    used_erp_inv = set()
    used_ven_inv = set()
    for e_idx, e in inv_erp.iterrows():
        if e_idx in used_erp_inv:
            continue
        e_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        e_date = e.get("date_erp", "")
        for v_idx, v in inv_ven.iterrows():
            if v_idx in used_ven_inv:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            v_date = v.get("date_ven", "")
            if e_inv == v_inv:
                diff = abs(e_amt - v_amt)
                if diff <= 0.01:
                    matched_perfect.append({
                        "ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt,
                        "Vendor Amount": v_amt, "Difference": round(diff, 2), "ERP Date": e_date,
                        "Vendor Date": v_date, "Status": "Perfect Match"
                    })
                    used_erp_inv.add(e_idx)
                    used_ven_inv.add(v_idx)
                    break
                elif diff < 1.0:
                    matched_diff.append({
                        "ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt,
                        "Vendor Amount": v_amt, "Difference": round(diff, 2), "ERP Date": e_date,
                        "Vendor Date": v_date, "Status": "Difference Match"
                    })
                    used_erp_inv.add(e_idx)
                    used_ven_inv.add(v_idx)
                    break
    remain_erp = inv_erp[~inv_erp.index.isin(used_erp_inv)]
    remain_ven = inv_ven[~inv_ven.index.isin(used_ven_inv)]
    for e_idx, e in remain_erp.iterrows():
        if e_idx in used_erp_inv:
            continue
        e_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        e_date = e.get("date_erp", "")
        if not e_date:
            continue
        best_match = None
        best_ratio = 0
        for v_idx, v in remain_ven.iterrows():
            if v_idx in used_ven_inv:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            v_date = v.get("date_ven", "")
            if not v_date or v_date != e_date:
                continue
            diff = abs(e_amt - v_amt)
            if diff > 0.01:
                continue
            ratio = fuzzy_ratio(e_inv, v_inv)
            if ratio > best_ratio and ratio >= 0.8:
                best_ratio = ratio
                best_match = v_idx
        if best_match is not None:
            v = remain_ven.loc[best_match]
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            v_date = v.get("date_ven", "")
            matched_tier3.append({
                "ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt,
                "Vendor Amount": v_amt, "Difference": 0.0, "ERP Date": e_date,
                "Vendor Date": v_date, "Fuzzy Ratio": round(best_ratio, 2), "Status": "Fuzzy Match"
            })
            used_erp_inv.add(e_idx)
            used_ven_inv.add(best_match)
    unmatch_erp = inv_erp[~inv_erp.index.isin(used_erp_inv)].rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp":"Date"})
    unmatch_ven = inv_ven[~inv_ven.index.isin(used_ven_inv)].rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven":"Date"})
    matched_pay = []
    used_erp_pay = set()
    used_ven_pay = set()
    for e_idx, e in pay_erp.iterrows():
        if e_idx in used_erp_pay:
            continue
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        e_date = e.get("date_erp", "")
        for v_idx, v in pay_ven.iterrows():
            if v_idx in used_ven_pay:
                continue
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            v_date = v.get("date_ven", "")
            diff = abs(e_amt - v_amt)
            if diff <= 0.01:
                matched_pay.append({
                    "ERP Amount": e_amt, "Vendor Amount": v_amt, "Difference": round(diff, 2),
                    "ERP Date": e_date, "Vendor Date": v_date
                })
                used_erp_pay.add(e_idx)
                used_ven_pay.add(v_idx)
                break
    unmatch_pay_erp = pay_erp[~pay_erp.index.isin(used_erp_pay)].rename(columns={"__amt": "Amount", "date_erp":"Date"})
    unmatch_pay_ven = pay_ven[~pay_ven.index.isin(used_ven_pay)].rename(columns={"__amt": "Amount", "date_ven":"Date"})
    return (pd.DataFrame(matched_perfect), pd.DataFrame(matched_diff), pd.DataFrame(matched_tier3),
            unmatch_erp[['Invoice','Amount','Date']], unmatch_ven[['Invoice','Amount','Date']],
            pd.DataFrame(matched_pay), unmatch_pay_erp[['Amount','Date']], unmatch_pay_ven[['Amount','Date']],
            pay_erp, pay_ven)
# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")
if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")
    for col in ["invoice_erp", "invoice_ven"]:
        if col in erp_df.columns:
            erp_df[col] = erp_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})
        if col in ven_df.columns:
            ven_df[col] = ven_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})
    with st.spinner("Reconciling..."):
        erp_use = consolidate_by_invoice(erp_df, "invoice_erp")
        ven_use = consolidate_by_invoice(ven_df, "invoice_ven")
        (tier1_perfect_df, tier1_diff_df, tier3_df, unmatch_erp_df, unmatch_ven_df,
         matched_pay_df, unmatch_pay_erp_df, unmatch_pay_ven_df, pay_erp, pay_ven) = match_invoices(erp_use, ven_use)
    st.success("✅ Reconciliation complete!")
    # ---------- TABS ----------
    tab1, tab2, tab3 = st.tabs(["📊 Summary", "🧾 Matches", "💰 Payments"])
    # --- SUMMARY TAB ---
    with tab1:
        col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
        col1.markdown(f"<div class='metric-box green'>Perfect<br><h2>{len(tier1_perfect_df)}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='metric-box orange'>Differences<br><h2>{len(tier1_diff_df)}</h2></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='metric-box teal'>Tier-2<br><h2>0</h2></div>", unsafe_allow_html=True)
        col4.markdown(f"<div class='metric-box purple'>Tier-3<br><h2>{len(tier3_df)}</h2></div>", unsafe_allow_html=True)
        col5.markdown(f"<div class='metric-box red'>Unmatched ERP<br><h2>{len(unmatch_erp_df)}</h2></div>", unsafe_allow_html=True)
        col6.markdown(f"<div class='metric-box pink'>Unmatched Vendor<br><h2>{len(unmatch_ven_df)}</h2></div>", unsafe_allow_html=True)
        col7.markdown(f"<div class='metric-box dark'>Matched Payments<br><h2>{len(matched_pay_df)}</h2></div>", unsafe_allow_html=True)
    # --- MATCHES TAB ---
    with tab2:
        st.markdown("### Tier-1 Perfect Matches")
        if not tier1_perfect_df.empty:
            st.dataframe(tier1_perfect_df, use_container_width=True)
        else:
            st.info("No perfect matches found.")
        st.markdown("### Tier-1 Difference Matches")
        if not tier1_diff_df.empty:
            st.dataframe(tier1_diff_df, use_container_width=True)
        else:
            st.info("No difference matches found.")
        st.markdown("### Tier-3 Fuzzy Matches")
        if not tier3_df.empty:
            st.dataframe(tier3_df, use_container_width=True)
        else:
            st.info("No fuzzy matches found.")
        st.markdown("### Missing in ERP (Present in Vendor)")
        st.dataframe(unmatch_ven_df, use_container_width=True)
        st.markdown("### Missing in Vendor (Present in ERP)")
        st.dataframe(unmatch_erp_df, use_container_width=True)
    # --- PAYMENTS TAB ---
    with tab3:
        total_erp_pay = pay_erp['__amt'].sum() if not pay_erp.empty else 0.0
        total_ven_pay = pay_ven['__amt'].sum() if not pay_ven.empty else 0.0
        pay_col1, pay_col2, pay_col3 = st.columns(3)
        pay_col1.markdown(f"<div class='metric-box green'>💸 Matched Payments<br><h2>{len(matched_pay_df)}</h2></div>", unsafe_allow_html=True)
        pay_col2.markdown(f"<div class='metric-box teal'>💰 Total ERP Payments<br><h2>{total_erp_pay:.2f}</h2></div>", unsafe_allow_html=True)
        pay_col3.markdown(f"<div class='metric-box purple'>💰 Total Vendor Payments<br><h2>{total_ven_pay:.2f}</h2></div>", unsafe_allow_html=True)
        st.markdown("### Matched Payments")
        if not matched_pay_df.empty:
            st.dataframe(matched_pay_df, use_container_width=True)
        else:
            st.info("No matched payments found.")
        st.markdown("### Unmatched ERP Payments")
        st.dataframe(unmatch_pay_erp_df, use_container_width=True)
        st.markdown("### Unmatched Vendor Payments")
        st.dataframe(unmatch_pay_ven_df, use_container_width=True)
        # Export all payments
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "ERP Payments"
        for r in dataframe_to_rows(pay_erp[['__amt', 'date_erp']], index=False, header=True):
            ws1.append(r)
        ws2 = wb.create_sheet("Vendor Payments")
        for r in dataframe_to_rows(pay_ven[['__amt', 'date_ven']], index=False, header=True):
            ws2.append(r)
        output = BytesIO()
        wb.save(output)
        st.download_button("📥 Export All Payments", data=output.getvalue(), file_name="all_payments.xlsx")
