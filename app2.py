import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher

# ==================== PAGE CONFIG & CSS ======================
st.set_page_config(page_title="ReconRaptor — Vendor Reconciliation", layout="wide")
st.markdown(
    """
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
    .perfect-match {background:#2E7D32;color:#fff;font-weight:bold;}
    .difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
    .tier2-match {background:#26A69A;color:#fff;font-weight:bold;}
    .tier3-match {background:#7E57C2;color:#fff;font-weight:bold;}
    .missing-erp {background:#C62828;color:#fff;font-weight:bold;}
    .missing-vendor {background:#AD1457;color:#fff;font-weight:bold;}
    .payment-match {background:#004D40;color:#fff;font-weight:bold;}
</style>
""",
    unsafe_allow_html=True,
)

# ==================== TITLES =========================
st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.3rem; color: #555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "": return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if s.count(",") == 1 and s.count(".") == 1:
        s = s.replace(".", "").replace(",", ".") if s.find(",") > s.find(".") else s.replace(",", "")
    elif s.count(",") == 1: s = s.replace(",", ".")
    elif s.count(".") > 1: s = s.replace(".", "", s.count(".") - 1)
    try: return float(s)
    except: return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "": return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    formats = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y", "%m-%d-%Y",
               "%Y/%m/%d", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
               "%m/%d/%y", "%m-%d-%y", "%Y.%m.%d"]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d): return d.strftime("%Y-%m-%d")
        except: continue
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(d): d = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s = str(v).strip().lower()
    parts = re.split(r"[-_]", s)
    for p in reversed(parts):
        if re.fullmatch(r"\d{1,}", p) and not re.fullmatch(r"20[0-3]\d", p):
            s = p.lstrip("0"); break
    s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|apd|vs)\W*", "", s)
    s = re.sub(r"20\d{2}", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    s = re.sub(r"[^\d]", "", s)
    return s or "0"

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "número", "document", "doc", "ref", "referencia",
                    "nº factura", "num factura", "alternative document", "document number", "αρ.", "αριθμός", "νουμερο",
                    "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου", "αριθμός τιμολογίου",
                    "αριθμός παραστατικού", "κωδικός τιμολογίου", "τιμολόγιο", "αρ. παραστατικού",
                    "παραστατικό τιμολογίου", "κωδικός παραστατικού"],
        "credit": ["credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito", "abono", "abonos",
                   "importe haber", "valor haber", "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού",
                   "ποσό πίστωσης", "πιστωτικό ποσό"],
        "debit": ["debit", "debe", "cargo", "importe", "importe total", "valor", "monto", "amount", "document value",
                  "charge", "total", "totale", "totales", "totals", "base imponible", "importe factura",
                  "importe neto", "χρέωση", "αξία", "αξία τιμολογίου", "ποσό χρέωσης", "συνολική αξία",
                  "καθαρή αξία", "ποσό", "ποσό τιμολογίου"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "descripción", "detalle", "detalles", "razon",
                   "razón", "observaciones", "comentario", "comentarios", "explicacion", "αιτιολογία", "περιγραφή",
                   "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή", "description", "περιγραφή τιμολογίου",
                   "αιτιολογία παραστατικού", "λεπτομέρειες"],
        "date": ["date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento", "ημερομηνία",
                 "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού", "issue date", "transaction date",
                 "emission date", "posting date", "ημερομηνία τιμολογίου", "ημερομηνία έκδοσης τιμολογίου",
                 "ημερομηνία καταχώρισης", "ημερ. έκδοσης", "ημερ. παραστατικού", "ημερομηνία έκδοσης παραστατικού"]
    }
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases): rename_map[col] = f"{key}_{tag}"
    out = df.rename(columns=rename_map)
    for req in ["debit", "credit"]:
        c = f"{req}_{tag}"
        if c not in out.columns: out[c] = 0.0
    if f"date_{tag}" in out.columns: out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

def style(df, css): return df.style.apply(lambda _: [css] * len(_), axis=1)

# ==================== CONSOLIDATION ==========================
def consolidate_by_invoice(df: pd.DataFrame, inv_col: str) -> pd.DataFrame:
    records = []
    if inv_col not in df.columns:
        return pd.DataFrame(columns=df.columns)
    tag = "erp" if "erp" in inv_col else "ven"
    debit_col = f"debit_{tag}"
    credit_col = f"credit_{tag}"
    for inv, group in df.groupby(inv_col, dropna=False):
        if group.empty: continue
        net = (
            group[debit_col].apply(normalize_number).sum() -
            group[credit_col].apply(normalize_number).sum()
        )
        net = round(net, 2)
        base = group.iloc[0].copy()
        base["__amt"] = abs(net)
        base["__type"] = "INV" if net >= 0 else "CN"
        base[debit_col] = max(net, 0.0)
        base[credit_col] = -min(net, 0.0)
        records.append(base)
    return pd.DataFrame(records).reset_index(drop=True)

# ==================== MATCHING CORE (TIER-1 FIXED) ==========================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor = set()

    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        pay_pat = [r"πληρωμ", r"payment", r"remittance", r"bank", r"transfer", r"trf", r"pago", r"pagado", r"εξοφλη"]
        if any(p in r for p in pay_pat):
            return "IGNORE"
        if any(k in r for k in ["credit", "nota", "abono", "cn", "πιστω", "πίστωση"]):
            return "CN"
        if any(k in r for k in ["factura", "invoice", "inv", "τιμολ", "παραστα"]) or normalize_number(row.get(f"debit_{tag}", 0)) > 0:
            return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)

    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)

    erp_use = consolidate_by_invoice(erp_df[erp_df["__type"] != "IGNORE"].copy(), "invoice_erp")
    ven_use = consolidate_by_invoice(ven_df[ven_df["__type"] != "IGNORE"].copy(), "invoice_ven")

    # 🔧 Normalize invoice codes to clean comparable strings
    erp_use["__inv_clean"] = erp_use["invoice_erp"].apply(lambda x: re.sub(r"\.0+$", "", str(clean_invoice_code(x)).strip()))
    ven_use["__inv_clean"] = ven_use["invoice_ven"].apply(lambda x: re.sub(r"\.0+$", "", str(clean_invoice_code(x)).strip()))

    for e_idx, e in erp_use.iterrows():
        e_code = e["__inv_clean"]
        e_amt = round(float(e["__amt"]), 2)
        e_typ = e.get("__type", "INV")

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor:
                continue
            v_code = v["__inv_clean"]
            v_amt = round(float(v["__amt"]), 2)
            v_typ = v.get("__type", "INV")

            # 🧠 strict but consistent comparison
            if e_typ != v_typ or e_code != v_code:
                continue

            diff = abs(e_amt - v_amt)
            if diff <= 0.01:
                status = "Perfect Match"
            elif diff < 1.00:
                status = "Difference Match"
            else:
                continue

            matched.append({
                "ERP Invoice": e.get("invoice_erp", ""),
                "Vendor Invoice": v.get("invoice_ven", ""),
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": round(diff, 2),
                "Status": status
            })
            used_vendor.add(v_idx)
            break

    matched_df = pd.DataFrame(matched)
    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_ven = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()

    miss_erp = erp_use[~erp_use["invoice_erp"].isin(matched_erp)].copy()
    miss_ven = ven_use[~ven_use["invoice_ven"].isin(matched_ven)].copy()

    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})

    keep = ["Invoice", "Amount", "Date"]
    miss_erp = miss_erp[[c for c in keep if c in miss_erp.columns]].reset_index(drop=True)
    miss_ven = miss_ven[[c for c in keep if c in miss_ven.columns]].reset_index(drop=True)

    return matched_df, miss_erp, miss_ven


# ==================== TIERS ==========================
def tier2_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty: return pd.DataFrame(), set(), set(), erp_miss.copy(), ven_miss.copy()
    e = erp_miss.copy(); v = ven_miss.copy(); matches, used_e, used_v = [], set(), set()
    for ei, er in e.iterrows():
        if ei in used_e: continue
        e_inv = str(er.get("Invoice", "")); e_amt = round(float(er.get("Amount", 0.0)), 2); e_code = clean_invoice_code(e_inv)
        for vi, vr in v.iterrows():
            if vi in used_v: continue
            v_inv = str(vr.get("Invoice", "")); v_amt = round(float(vr.get("Amount", 0.0)), 2); v_code = clean_invoice_code(v_inv)
            diff = abs(e_amt - v_amt); sim = fuzzy_ratio(e_code, v_code)
            if diff <= 1.00 and sim >= 0.85:
                matches.append({"ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt,
                                "Vendor Amount": v_amt, "Difference": round(diff, 2), "Fuzzy Score": round(sim, 2), "Match Type": "Tier-2"})
                used_e.add(ei); used_v.add(vi); break
    mdf = pd.DataFrame(matches)
    rem_e = e[~e.index.isin(used_e)].copy(); rem_v = v[~v.index.isin(used_v)].copy()
    return mdf, used_e, used_v, rem_e, rem_v

def tier3_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty: return pd.DataFrame(), set(), set(), erp_miss.copy(), ven_miss.copy()
    e = erp_miss.copy(); v = ven_miss.copy(); matches, used_e, used_v = [], set(), set()
    for ei, er in e.iterrows():
        if ei in used_e: continue
        e_inv = str(er.get("Invoice", "")); e_amt = round(float(er.get("Amount", 0.0)), 2)
        e_date = normalize_date(er.get("Date", "")) if "Date" in er else ""; e_code = clean_invoice_code(e_inv)
        if not e_date: continue
        for vi, vr in v.iterrows():
            if vi in used_v: continue
            v_inv = str(vr.get("Invoice", "")); v_amt = round(float(vr.get("Amount", 0.0)), 2)
            v_date = normalize_date(vr.get("Date", "")) if "Date" in vr else ""; v_code = clean_invoice_code(v_inv)
            if not v_date: continue
            sim = fuzzy_ratio(e_code, v_code)
            if e_date == v_date and sim >= 0.90:
                diff = abs(e_amt - v_amt)
                matches.append({"ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt,
                                "Vendor Amount": v_amt, "Difference": round(diff, 2), "Fuzzy Score": round(sim, 2),
                                "Date": e_date, "Match Type": "Tier-3"})
                used_e.add(ei); used_v.add(vi); break
    mdf = pd.DataFrame(matches)
    rem_e = e[~e.index.isin(used_e)].copy(); rem_v = v[~v.index.isin(used_v)].copy()
    return mdf, used_e, used_v, rem_e, rem_v

# ==================== PAYMENTS ==========================
def extract_payments(erp_df, ven_df):
    pay_kw = ["πληρωμή", "payment", "payment remittance", "remittance", "bank transfer", "transferencia",
              "trf", "remesa", "pago", "deposit", "μεταφορά", "έμβασμα", "εξοφληση", "pagado", "paid", "bank"]
    excl_kw = ["invoice of expenses", "expense invoice", "τιμολόγιο εξόδων", "διόρθωση",
               "correction", "reclass", "adjustment", "μεταφορά υπολοίπου"]
    def is_pay(row, tag):
        txt = str(row.get(f"reason_{tag}", "")).lower()
        return any(k in txt for k in pay_kw) and not any(b in txt for b in excl_kw) \
               and ((tag == "erp" and normalize_number(row.get("debit_erp", 0)) > 0) or
                    (tag == "ven" and normalize_number(row.get("credit_ven", 0)) > 0))
    erp_pay = erp_df[erp_df.apply(lambda r: is_pay(r, "erp"), axis=1)].copy() if "reason_erp" in erp_df.columns else pd.DataFrame()
    ven_pay = ven_df[ven_df.apply(lambda r: is_pay(r, "ven"), axis=1)].copy() if "reason_ven" in ven_df.columns else pd.DataFrame()
    if not erp_pay.empty: erp_pay["Amount"] = erp_pay.apply(lambda r: abs(normalize_number(r["debit_erp"]) - normalize_number(r["credit_erp"])), axis=1)
    if not ven_pay.empty: ven_pay["Amount"] = ven_pay.apply(lambda r: abs(normalize_number(r["debit_ven"]) - normalize_number(r["credit_ven"])), axis=1)
    matched = []; used = set()
    for _, e in erp_pay.iterrows():
        for vi, v in ven_pay.iterrows():
            if vi in used: continue
            if abs(e["Amount"] - v["Amount"]) <= 0.01:
                matched.append({"ERP Reason": e.get("reason_erp", ""), "Vendor Reason": v.get("reason_ven", ""),
                                "ERP Amount": round(e["Amount"], 2), "Vendor Amount": round(v["Amount"], 2),
                                "Difference": round(abs(e["Amount"] - v["Amount"]), 2)})
                used.add(vi); break
    return erp_pay, ven_pay, pd.DataFrame(matched)

# ==================== EXCEL EXPORT =========================
def export_excel(t1, t2, t3, miss_erp, miss_ven, pay_match):
    wb = Workbook()
    def hdr(ws, row, color):
        for c in ws[row]:
            c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
    ws1 = wb.active; ws1.title = "Tier1"
    if not t1.empty:
        for r in dataframe_to_rows(t1, index=False, header=True): ws1.append(r)
        hdr(ws1, 1, "1E88E5")
    ws2 = wb.create_sheet("Tier2")
    if not t2.empty:
        for r in dataframe_to_rows(t2, index=False, header=True): ws2.append(r)
        hdr(ws2, 1, "26A69A")
    ws3 = wb.create_sheet("Tier3")
    if not t3.empty:
        for r in dataframe_to_rows(t3, index=False, header=True): ws3.append(r)
        hdr(ws3, 1, "7E57C2")
    ws4 = wb.create_sheet("Missing")
    cur = 1
    if not miss_ven.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max(3, miss_ven.shape[1]))
        ws4.cell(cur, 1, "Missing in ERP").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_ven, index=False, header=True): ws4.append(r)
        hdr(ws4, cur, "C62828"); cur = ws4.max_row + 3
    if not miss_erp.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max(3, miss_erp.shape[1]))
        ws4.cell(cur, 1, "Missing in Vendor").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_erp, index=False, header=True): ws4.append(r)
        hdr(ws4, cur, "AD1457")
    ws5 = wb.create_sheet("Payments")
    if not pay_match.empty:
        for r in dataframe_to_rows(pay_match, index=False, header=True): ws5.append(r)
        hdr(ws5, 1, "004D40")
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3
    buf = BytesIO(); wb.save(buf); buf.seek(0); return buf

# ==================== UI =========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")

        # --- NORMALISE INVOICE NUMBERS ---
        for col in ["invoice_erp", "invoice_ven"]:
            if col in erp_df.columns:
                erp_df[col] = erp_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})
            if col in ven_df.columns:
                ven_df[col] = ven_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})

        with st.spinner("Analyzing..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)
            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)
            erp_pay, ven_pay, pay_match = extract_payments(erp_df, ven_df)

        st.success("Complete!")

        # ---------- METRICS ----------
        st.markdown('<h2 class="section-title">Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()
        def safe_sum(df, col): return float(df[col].sum()) if not df.empty and col in df.columns else 0.0

        with c1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("Perfect", len(perf))
            st.markdown(f"**Diff:** {safe_sum(perf, 'Difference'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
            st.metric("Differences", len(diff))
            st.markdown(f"**Diff:** {safe_sum(diff, 'Difference'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c3:
            st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
            st.metric("Tier-2", len(tier2))
            st.markdown(f"**Diff:** {safe_sum(tier2, 'Difference'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="metric-container tier3-match">', unsafe_allow_html=True)
            st.metric("Tier-3", len(tier3))
            st.markdown(f"**Diff:** {safe_sum(tier3, 'Difference'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c5:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("Unmatched ERP", len(final_erp_miss))
            st.markdown(f"**Total:** {final_erp_miss['Amount'].sum():,.2f}" if "Amount" in final_erp_miss else "**Total:** 0.00", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c6:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("Unmatched Vendor", len(final_ven_miss))
            st.markdown(f"**Total:** {final_ven_miss['Amount'].sum():,.2f}" if "Amount" in final_ven_miss else "**Total:** 0.00", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c7:
            st.markdown('<div class="metric-container payment-match">', unsafe_allow_html=True)
            st.metric("Payment Matches", len(pay_match))
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")

        # ---------- DISPLAY ----------
        st.markdown('<h2 class="section-title">Tier-1: Exact Matches</h2>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Perfect Matches**")
            if not perf.empty:
                st.dataframe(style(perf[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference']], "background:#2E7D32;color:#fff;font-weight:bold;"), use_container_width=True)
            else:
                st.info("No perfect matches")
        with col_b:
            st.markdown("**Amount Differences (< 1.00)**")
            if not diff.empty:
                st.dataframe(style(diff[['ERP Invoice', 'Vendor Invoice', 'ERP Amount', 'Vendor Amount', 'Difference']], "background:#FF8F00;color:#fff;font-weight:bold;"), use_container_width=True)
            else:
                st.success("No differences")

        st.markdown('<h2 class="section-title">Tier-2: Fuzzy + Small Diff</h2>', unsafe_allow_html=True)
        if not tier2.empty:
            st.dataframe(style(tier2, "background:#26A69A;color:#fff;font-weight:bold;"), use_container_width=True)
        else:
            st.info("No Tier-2 matches")

        st.markdown('<h2 class="section-title">Tier-3: Date + Strict Fuzzy</h2>', unsafe_allow_html=True)
        if not tier3.empty:
            st.dataframe(style(tier3, "background:#7E57C2;color:#fff;font-weight:bold;"), use_container_width=True)
        else:
            st.info("No Tier-3 matches")

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown('<h2 class="section-title">Missing in ERP</h2>', unsafe_allow_html=True)
            if not final_ven_miss.empty:
                st.dataframe(style(final_ven_miss, "background:#AD1457;color:#fff;font-weight:bold;"), use_container_width=True)
            else:
                st.success("All vendor invoices found in ERP")
        with col_m2:
            st.markdown('<h2 class="section-title">Missing in Vendor</h2>', unsafe_allow_html=True)
            if not final_erp_miss.empty:
                st.dataframe(style(final_erp_miss, "background:#C62828;color:#fff;font-weight:bold;"), use_container_width=True)
            else:
                st.success("All ERP invoices found in vendor")

        st.markdown('<h2 class="section-title">Payments</h2>', unsafe_allow_html=True)
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("**ERP Payments**")
            if not erp_pay.empty:
                disp = erp_pay[['reason_erp', 'debit_erp', 'credit_erp', 'Amount']].copy()
                disp.columns = ['Reason', 'Debit', 'Credit', 'Net']
                st.dataframe(disp.style.apply(lambda _: ['background:#4CAF50;color:#fff'] * len(_), axis=1), use_container_width=True)
            else:
                st.info("No ERP payments")
        with col_p2:
            st.markdown("**Vendor Payments**")
            if not ven_pay.empty:
                disp = ven_pay[['reason_ven', 'debit_ven', 'credit_ven', 'Amount']].copy()
                disp.columns = ['Reason', 'Debit', 'Credit', 'Net']
                st.dataframe(disp.style.apply(lambda _: ['background:#2196F3;color:#fff'] * len(_), axis=1), use_container_width=True)
            else:
                st.info("No vendor payments")

        if not pay_match.empty:
            st.markdown("**Matched Payments**")
            st.dataframe(pay_match.style.apply(lambda _: ['background:#004D40;color:#fff;font-weight:bold'] * len(_), axis=1), use_container_width=True)

        # ---------- EXPORT ----------
        st.markdown('<h2 class="section-title">Download Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(tier1, tier2, tier3, final_erp_miss, final_ven_miss, pay_match)
        st.download_button("Download Excel", data=excel_buf, file_name="ReconRaptor_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Need columns: invoice, debit/credit, date, reason")
