# --------------------------------------------------------------
# ReconRaptor – Vendor Reconciliation (FINAL, Smart Consolidation)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
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
    if pd.isna(v) or str(v).strip() == "":
        return 0.0
    s = re.sub(r"[^\d,.-]", "", str(v).strip())
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
    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%m-%d-%Y",
        "%Y/%m/%d", "%Y-%m-%d",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
        "%m/%d/%y", "%m-%d-%y",
        "%Y.%m.%d",
    ]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d):
                return d.strftime("%Y-%m-%d")
        except:
            continue
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(d):
        d = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v:
        return ""
    s = str(v).strip().lower()
    parts = re.split(r"[-*]", s)
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
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "document", "doc", "ref", "referencia",
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
    return df.style.apply(lambda x: [css] * len(x), axis=1)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor = set()

    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        pay_pat = [
            r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^remittance",
            r"^bank\s*transfer", r"^trf", r"^remesa", r"^pago",
            r"^εξοφληση", r"^paid"
        ]
        if any(re.search(p, r) for p in pay_pat):
            return "IGNORE"
        if any(k in r for k in ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]):
            return "CN"
        if any(k in r for k in ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]) or debit > 0:
            return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)
    erp_df["__amt"] = erp_df.apply(
        lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1
    )
    ven_df["__amt"] = ven_df.apply(
        lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1
    )

    erp_use = erp_df[erp_df["__type"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__type"] != "IGNORE"].copy()

    def consolidate_by_invoice(df, inv_col):
        if inv_col not in df.columns:
            return pd.DataFrame(columns=df.columns)
        records = []
        cancel_kw = [
            "cancel", "cancellation", "correct", "correction", "storno",
            "reversal", "void", "αντιλογισ", "ακυρω", "διόρθωση"
        ]
        for inv, group in df.groupby(inv_col, dropna=False):
            if group.empty:
                continue
            total = 0.0
            for _, row in group.iterrows():
                amt = normalize_number(row.get("__amt", 0))
                reason = str(row.get("reason_erp", "")) + " " + str(row.get("reason_ven", ""))
                reason = reason.lower()
                if any(k in reason for k in cancel_kw):
                    total -= amt
                elif row.get("__type", "INV") == "CN":
                    total -= amt
                else:
                    total += amt
            net = round(total, 2)
            if abs(net) < 0.01:
                continue
            base = group.iloc[0].copy()
            base["__amt"] = abs(net)
            base["__type"] = "INV" if net > 0 else "CN"
            records.append(base)
        return pd.DataFrame(records).reset_index(drop=True)

    erp_use = consolidate_by_invoice(erp_use, "invoice_erp")
    ven_use = consolidate_by_invoice(ven_use, "invoice_ven")

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        e_typ = e.get("__type", "INV")
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            v_typ = v.get("__type", "INV")
            if e_typ != v_typ or e_inv != v_inv:
                continue
            diff = abs(e_amt - v_amt)
            status = "Perfect Match" if diff <= 0.01 else "Difference Match"
            matched.append({
                "ERP Invoice": e_inv,
                "Vendor Invoice": v_inv,
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
    miss_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)].copy()
    miss_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp)].copy()
    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
    keep_cols = ["Invoice", "Amount", "Date"]
    miss_erp = miss_erp[[c for c in keep_cols if c in miss_erp.columns]].reset_index(drop=True)
    miss_ven = miss_ven[[c for c in keep_cols if c in miss_ven.columns]].reset_index(drop=True)
    return matched_df, miss_erp, miss_ven

# ------- Tier-2: fuzzy invoice code + small amount tolerance -------
def tier2_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty:
        return pd.DataFrame(), set(), set(), erp_miss.copy(), ven_miss.copy()
    e = erp_miss.copy()
    v = ven_miss.copy()
    matches, used_e, used_v = [], set(), set()
    for ei, er in e.iterrows():
        if ei in used_e:
            continue
        e_inv = str(er.get("Invoice", ""))
        e_amt = round(float(er.get("Amount", 0.0)), 2)
        e_code = clean_invoice_code(e_inv)
        for vi, vr in v.iterrows():
            if vi in used_v:
                continue
            v_inv = str(vr.get("Invoice", ""))
            v_amt = round(float(vr.get("Amount", 0.0)), 2)
            v_code = clean_invoice_code(v_inv)
            diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_code, v_code)
            if diff <= 1.00 and sim >= 0.85:
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": round(diff, 2),
                    "Fuzzy Score": round(sim, 2),
                    "Match Type": "Tier-2"
                })
                used_e.add(ei)
                used_v.add(vi)
                break
    mdf = pd.DataFrame(matches)
    rem_e = e[~e.index.isin(used_e)].copy()
    rem_v = v[~v.index.isin(used_v)].copy()
    return mdf, used_e, used_v, rem_e, rem_v

# ------- Tier-3: same DATE + strong fuzzy invoice (no amount threshold) -------
def tier3_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty:
        return pd.DataFrame(), set(), set(), erp_miss.copy(), ven_miss.copy()
    e = erp_miss.copy()
    v = ven_miss.copy()
    matches, used_e, used_v = [], set(), set()
    for ei, er in e.iterrows():
        if ei in used_e:
            continue
        e_inv = str(er.get("Invoice", ""))
        e_amt = round(float(er.get("Amount", 0.0)), 2)
        e_date = normalize_date(er.get("Date", "")) if "Date" in er else ""
        e_code = clean_invoice_code(e_inv)
        if not e_date:
            continue
        for vi, vr in v.iterrows():
            if vi in used_v:
                continue
            v_inv = str(vr.get("Invoice", ""))
            v_amt = round(float(vr.get("Amount", 0.0)), 2)
            v_date = normalize_date(vr.get("Date", "")) if "Date" in vr else ""
            v_code = clean_invoice_code(v_inv)
            if not v_date:
                continue
            sim = fuzzy_ratio(e_code, v_code)
            if e_date == v_date and sim >= 0.75:
                diff = abs(e_amt - v_amt)
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": round(diff, 2),
                    "Fuzzy Score": round(sim, 2),
                    "Date": e_date,
                    "Match Type": "Tier-3"
                })
                used_e.add(ei)
                used_v.add(vi)
                break
    mdf = pd.DataFrame(matches)
    rem_e = e[~e.index.isin(used_e)].copy()
    rem_v = v[~v.index.isin(used_v)].copy()
    return mdf, used_e, used_v, rem_e, rem_v

# ------- Payments detection & matching -------
def extract_payments(erp_df, ven_df):
    pay_kw = [
        "πληρωμή", "payment", "payment remittance", "remittance", "bank transfer",
        "transferencia", "trf", "remesa", "pago", "deposit", "μεταφορά", "έμβασμα",
        "εξοφληση", "pagado", "paid"
    ]
    excl_kw = [
        "invoice of expenses", "expense invoice", "τιμολόγιο εξόδων", "διόρθωση",
        "correction", "reclass", "adjustment", "μεταφορά υπολοίπου"
    ]
    def is_pay(row, tag):
        txt = str(row.get(f"reason_{tag}", "")).lower()
        return any(k in txt for k in pay_kw) and not any(b in txt for b in excl_kw) \
               and ((tag == "erp" and normalize_number(row.get("debit_erp", 0)) > 0) or
                    (tag == "ven" and normalize_number(row.get("credit_ven", 0)) > 0))
    erp_pay = erp_df[erp_df.apply(lambda r: is_pay(r, "erp"), axis=1)].copy() if "reason_erp" in erp_df.columns else pd.DataFrame()
    ven_pay = ven_df[ven_df.apply(lambda r: is_pay(r, "ven"), axis=1)].copy() if "reason_ven" in ven_df.columns else pd.DataFrame()
    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(lambda r: abs(normalize_number(r["debit_erp"]) - normalize_number(r["credit_erp"])), axis=1)
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(lambda r: abs(normalize_number(r["debit_ven"]) - normalize_number(r["credit_ven"])), axis=1)
    matched = []
    used = set()
    for _, e in erp_pay.iterrows():
        for vi, v in ven_pay.iterrows():
            if vi in used:
                continue
            if abs(e["Amount"] - v["Amount"]) <= 0.01:
                matched.append({
                    "ERP Reason": e.get("reason_erp", ""),
                    "Vendor Reason": v.get("reason_ven", ""),
                    "ERP Amount": round(e["Amount"], 2),
                    "Vendor Amount": round(v["Amount"], 2),
                    "Difference": round(abs(e["Amount"] - v["Amount"]), 2)
                })
                used.add(vi)
                break
    return erp_pay, ven_pay, pd.DataFrame(matched)

# ==================== EXCEL EXPORT =========================
def export_excel(t1, t2, t3, miss_erp, miss_ven, pay_match):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not t1.empty:
            t1.to_excel(writer, sheet_name='Tier1', index=False)
        if not t2.empty:
            t2.to_excel(writer, sheet_name='Tier2', index=False)
        if not t3.empty:
            t3.to_excel(writer, sheet_name='Tier3', index=False)
        if not miss_ven.empty or not miss_erp.empty:
            miss_combined = pd.concat([
                miss_ven.assign(Source="Missing in ERP") if not miss_ven.empty else pd.DataFrame(),
                miss_erp.assign(Source="Missing in Vendor") if not miss_erp.empty else pd.DataFrame()
            ], ignore_index=True)
            miss_combined.to_excel(writer, sheet_name='Missing', index=False)
        if not pay_match.empty:
            pay_match.to_excel(writer, sheet_name='Payments', index=False)

        workbook = writer.book
        header_fill = PatternFill(start_color="FF6B35", end_color="FF6B35", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column].width = adjusted_width

    output.seek(0)
    return output

# ==================== UI =========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

# Initialize variables outside try
tier1 = miss_erp = miss_ven = tier2 = tier3 = final_erp_miss = final_ven_miss = erp_pay = ven_pay = pay_match = pd.DataFrame()

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")
        with st.spinner("Analyzing invoices..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)
            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)
            erp_pay, ven_pay, pay_match = extract_payments(erp_df, ven_df)
        st.success("Reconciliation Complete!")

        # [ALL METRICS & DISPLAY CODE — UNCHANGED — GOES HERE]
        # ... (your full metrics and display code)

    except Exception as e:
        st.error(f"Error during processing: {e}")
        st.info("Please check your files and try again.")

# ---------- EXPORT ----------
st.markdown('<h2 style="color: #FF6B35;">Download Missing Items</h2>', unsafe_allow_html=True)

if final_erp_miss.empty and final_ven_miss.empty:
    st.warning("No missing items to export!")
else:
    combined = pd.concat([
        final_erp_miss.assign(Source="Missing in ERP") if not final_erp_miss.empty else pd.DataFrame(),
        final_ven_miss.assign(Source="Missing in Vendor") if not final_ven_miss.empty else pd.DataFrame()
    ], ignore_index=True)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        combined.to_excel(writer, sheet_name="Missing_Only", index=False)
    buf.seek(0)

    st.download_button(
        label="Download Missing Items (Excel)",
        data=buf.getvalue(),
        file_name="Missing_Items.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
