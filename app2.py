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
# CONFIGURATION
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ======================================
# HELPERS
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
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(d):
            d = pd.to_datetime(s, errors="coerce", dayfirst=False)
        if pd.isna(d):
            return ""
        return d.strftime("%Y-%m-%d")
    except:
        return ""

def normalize_columns(df, tag):
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document",
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού",
            "issue date", "transaction date", "emission date", "posting date"
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
    # Normalize dates in date_erp or date_ven columns if they exist
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()
    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        payment_patterns = [
            r"^πληρωμ",
            r"^απόδειξη\s*πληρωμ",
            r"^payment",
            r"^bank\s*transfer",
            r"^trf",
            r"^remesa",
            r"^pago",
            r"^pagado",
            r"^transferencia",
            r"^εξοφληση",
            r"^paid",
        ]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό", "ακυρωτικό παραστατικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in credit_words):
            return "CN"
        elif any(k in reason for k in invoice_words) or credit > 0:
            return "INV"
        return "UNKNOWN"
    def calc_erp_amount(row):
        doc = row.get("__doctype", "")
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if doc == "INV":
            return abs(credit)
        elif doc == "CN":
            return -abs(charge if charge > 0 else credit)
        return 0.0
    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        payment_patterns = [
            r"^πληρωμ",
            r"^απόδειξη\s*πληρωμ",
            r"^payment",
            r"^bank\s*transfer",
            r"^trf",
            r"^remesa",
            r"^pago",
            r"^pagado",
            r"^transferencia",
            r"^εξοφληση",
            r"^paid",
        ]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό", "ακυρωτικό παραστατικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in credit_words) or credit > 0:
            return "CN"
        elif any(k in reason for k in invoice_words) or debit > 0:
            return "INV"
        return "UNKNOWN"
    def calc_vendor_amount(row):
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        doc = row.get("__doctype", "")
        if doc == "INV":
            return abs(debit)
        elif doc == "CN":
            return -abs(credit if credit > 0 else debit)
        return 0.0
    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)
    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()
    merged_rows = []
    for inv, group in erp_use.groupby("invoice_erp", dropna=False):
        if group.empty:
            continue
        if len(group) >= 3:
            group = group.tail(1)
        inv_rows = group[group["__doctype"] == "INV"]
        cn_rows = group[group["__doctype"] == "CN"]
        if not inv_rows.empty and not cn_rows.empty:
            total_inv = inv_rows["__amt"].sum()
            total_cn = cn_rows["__amt"].sum()
            net = round(total_inv + total_cn, 2)
            base_row = inv_rows.iloc[-1].copy()
            base_row["__amt"] = net
            merged_rows.append(base_row)
        else:
            merged_rows.append(group.iloc[-1])
    erp_use = pd.DataFrame(merged_rows).reset_index(drop=True)
    erp_use["__amt"] = erp_use["__amt"].astype(float)
    ven_use["__amt"] = ven_use["__amt"].astype(float)
    erp_use = erp_use.groupby(["invoice_erp", "__doctype"], as_index=False)["__amt"].sum()
    ven_use = ven_use.groupby(["invoice_ven", "__doctype"], as_index=False)["__amt"].sum()
    def extract_digits(v):
        return re.sub(r"\D", "", str(v or "")).lstrip("0")
    def clean_invoice_code(v):
        if not v:
            return ""
        s = str(v).strip().lower()
        parts = re.split(r"[-_]", s)
        for p in reversed(parts):
            if re.fullmatch(r"\d{4,}", p) and not re.fullmatch(r"20[0-3]\d", p):
                s = p.lstrip("0")
                break
        s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no)\W*", "", s)
        s = re.sub(r"20\d{2}", "", s)
        s = re.sub(r"[^a-z0-9]", "", s)
        s = re.sub(r"^0+", "", s)
        s = re.sub(r"[^\d]", "", s)
        return s
    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_code = clean_invoice_code(e_inv)
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_code = clean_invoice_code(v_inv)
            diff = round(e_amt - v_amt, 2)
            amt_close = abs(diff) < 0.05
            same_full = (e_inv == v_inv)
            same_clean = (e_code == v_code)
            same_type = (e["__doctype"] == v["__doctype"])
            e_num = re.sub(r".*?(\d{2,})$", r"\1", str(e_inv))
            v_num = re.sub(r".*?(\d{2,})$", r"\1", str(v_inv))
            if same_type and same_full:
                take_it = True
            elif same_type and same_clean and amt_close:
                take_it = True
            elif same_type and e_num == v_num and amt_close:
                take_it = True
            elif same_type and e_num == v_num and (amt_close or abs(diff) < 0.1):
                take_it = True
            else:
                take_it = False
            if take_it:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match" if amt_close else "Difference"
                })
                used_vendor_rows.add(v_idx)
                break
    matched_df = pd.DataFrame(matched)
    matched_erp = {m["ERP Invoice"] for _, m in matched_df.iterrows()}
    matched_ven = {m["Vendor Invoice"] for _, m in matched_df.iterrows()}
    # Conditionally include date columns if they exist
    erp_columns = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in erp_use.columns else [])
    ven_columns = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in ven_use.columns else [])
    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][ven_columns] \
        if "invoice_ven" in ven_use else pd.DataFrame()
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][erp_columns] \
        if "invoice_erp" in erp_use else pd.DataFrame()
    missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    return matched_df, missing_in_erp, missing_in_vendor

# ======================================
# TIER-2 MATCHING
# ======================================
def fuzzy_ratio(a, b):
    """Calculate similarity ratio between two strings."""
    return SequenceMatcher(None, str(a), str(b)).ratio()

def tier2_match(erp_missing, ven_missing):
    """Perform Tier-2 matching on unmatched invoices using fuzzy matching, date, and amount."""
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), erp_missing.copy(), ven_missing.copy()
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt", "Date": "date_erp"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt", "Date": "date_ven"}).copy()
    e_df["date_norm"] = e_df["date_erp"].apply(normalize_date) if "date_erp" in e_df.columns else ""
    v_df["date_norm"] = v_df["date_ven"].apply(normalize_date) if "date_ven" in v_df.columns else ""
    matches, used_e, used_v = [], set(), set()
    for e_idx, e in e_df.iterrows():
        if e_idx in used_e:
            continue
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0)), 2)
        e_date = e.get("date_norm", "")
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0)), 2)
            v_date = v.get("date_norm", "")
            diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_inv, v_inv)
            # Match if amounts are close, fuzzy score is high, and dates match
            if diff < 0.05 and sim >= 0.8 and e_date == v_date and e_date != "" and v_date != "":
                matches.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Fuzzy Score": round(sim, 2),
                    "Date": e_date,
                    "Match Type": "Tier-2"
                })
                used_e.add(e_idx)
                used_v.add(v_idx)
                break
    tier2_matches = pd.DataFrame(matches)
    # Conditionally include date columns in remaining unmatched
    erp_columns = ["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in e_df.columns else [])
    ven_columns = ["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in v_df.columns else [])
    remaining_erp_missing = e_df[~e_df.index.isin(used_e)][erp_columns].rename(
        columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"}
    )
    remaining_ven_missing = v_df[~v_df.index.isin(used_v)][ven_columns].rename(
        columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"}
    )
    return tier2_matches, remaining_erp_missing, remaining_ven_missing

# ======================================
def extract_payments(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    """Extract and match payment transactions."""
    payment_keywords = [
        "πληρωμή", "payment", "bank transfer", "transferencia bancaria",
        "transfer", "trf", "remesa", "pago", "deposit", "μεταφορά", "έμβασμα",
        "εξοφληση", "pagado", "paid"
    ]
    exclude_keywords = [
        "τιμολόγιο", "invoice", "παραστατικό", "έξοδα", "expenses", "expense",
        "invoice of expenses", "expense invoice", "τιμολόγιο εξόδων",
        "διόρθωση", "διορθώσεις", "correction", "reclass", "adjustment",
        "μεταφορά υπολοίπου", "balance transfer"
    ]
    def is_real_payment(row: pd.Series, tag: str) -> bool:
        """Returns True if the row is a payment, considering keywords and debit/credit conditions."""
        text = str(row.get(f"reason_{tag}", "")).lower()
        has_payment = any(k in text for k in payment_keywords)
        has_exclusion = any(bad in text for bad in exclude_keywords)
        if tag == "erp":
            debit = normalize_number(row.get("debit_erp", 0))
            return has_payment and not has_exclusion and debit > 0
        elif tag == "ven":
            credit = normalize_number(row.get("credit_ven", 0))
            return has_payment and not has_exclusion and credit > 0
        return False
    # --- Filter ERP & Vendor payments ---
    erp_pay = (
        erp_df[erp_df.apply(lambda x: is_real_payment(x, "erp"), axis=1)].copy()
        if "reason_erp" in erp_df else pd.DataFrame()
    )
    ven_pay = (
        ven_df[ven_df.apply(lambda x: is_real_payment(x, "ven"), axis=1)].copy()
        if "reason_ven" in ven_df else pd.DataFrame()
    )
    # --- Calculate amounts ---
    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_erp")) - normalize_number(r.get("credit_erp"))),
            axis=1
        )
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_ven")) - normalize_number(r.get("credit_ven"))),
            axis=1
        )
    # --- Match payments ---
    matched_payments = []
    used_vendor = set()
    for _, e in erp_pay.iterrows():
        for v_idx, v in ven_pay.iterrows():
            if v_idx in used_vendor:
                continue
            diff = abs(e["Amount"] - v["Amount"])
            if diff < 0.05:
                matched_payments.append({
                    "ERP Reason": e.get("reason_erp", ""),
                    "Vendor Reason": v.get("reason_ven", ""),
                    "ERP Amount": round(float(e["Amount"]), 2),
                    "Vendor Amount": round(float(v["Amount"]), 2),
                    "Difference": round(diff, 2)
                })
                used_vendor.add(v_idx)
                break
    return erp_pay, ven_pay, pd.DataFrame(matched_payments)

# ======================================
def export_reconciliation_excel(matched, erp_missing, ven_missing, matched_pay, tier2_matches):
    """Export reconciliation results to Excel with four tabs."""
    wb = Workbook()
    
    # Helper function to style headers only
    def style_header(ws, row, color):
        for cell in ws[row]:
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True, size=12)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # ===== Sheet 1: Matches (Tier-1) =====
    ws1 = wb.active
    ws1.title = "Matches"
    if not matched.empty:
        for r in dataframe_to_rows(matched, index=False, header=True):
            ws1.append(r)
        style_header(ws1, 1, "1E88E5")  # Blue header for Matches
    
    # ===== Sheet 2: Missing =====
    ws2 = wb.create_sheet("Missing")
    current_row = 1
    
    # --- Missing in ERP ---
    if not erp_missing.empty:
        ws2.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=erp_missing.shape[1])
        ws2.cell(row=current_row, column=1, value="Missing in ERP (found in vendor but not in ERP)").font = Font(bold=True, size=14, color="000000")
        current_row += 2
        for r in dataframe_to_rows(erp_missing, index=False, header=True):
            ws2.append(r)
        style_header(ws2, current_row, "C62828")  # Red header for Missing in ERP
        current_row = ws2.max_row + 3
    
    # --- Missing in Vendor ---
    if not ven_missing.empty:
        ws2.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=ven_missing.shape[1])
        ws2.cell(row=current_row, column=1, value="Missing in Vendor (found in ERP but not in vendor)").font = Font(bold=True, size=14, color="000000")
        current_row += 2
        for r in dataframe_to_rows(ven_missing, index=False, header=True):
            ws2.append(r)
        style_header(ws2, current_row, "AD1457")  # Pink header for Missing in Vendor
        current_row = ws2.max_row + 3
    
    # ===== Sheet 3: Payments =====
    ws3 = wb.create_sheet("Payments")
    if not matched_pay.empty:
        for r in dataframe_to_rows(matched_pay, index=False, header=True):
            ws3.append(r)
        style_header(ws3, 1, "2E7D32")  # Green header for Payments
    
    # ===== Sheet 4: Tier-2 Matches =====
    ws4 = wb.create_sheet("Tier-2 Matches")
    if not tier2_matches.empty:
        for r in dataframe_to_rows(tier2_matches, index=False, header=True):
            ws4.append(r)
        style_header(ws4, 1, "26A69A")  # Teal header for Tier-2 Matches
    
    # ===== Auto-fit columns for all sheets =====
    for ws in [ws1, ws2, ws3, ws4]:
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3
    
    # ===== Save to buffer =====
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ======================================
# STREAMLIT UI
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])
if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")
    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)
        tier2_matches, remaining_erp_missing, remaining_ven_missing = tier2_match(erp_missing, ven_missing)
    st.success("✅ Reconciliation complete")
    # ====== HIGHLIGHTING ======
    def highlight_row(row):
        if row.get("Status") == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row.get("Status") == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        elif row.get("Match Type") == "Tier-2":
            return ['background-color: #26A69A; color: white'] * len(row)
        return [''] * len(row)
    # ====== MATCHED (TIER-1) ======
    st.subheader("📊 Tier-1 Matches / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No Tier-1 matches found.")
    # ====== TIER-2 MATCHES ======
    st.subheader("📈 Tier-2 Matches (Fuzzy Matching)")
    if not tier2_matches.empty:
        st.dataframe(tier2_matches.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No Tier-2 matches found.")
    # ====== MISSING ======
    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(
            erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True
        )
    else:
        st.success("✅ No missing invoices in ERP.")
    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(
            ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True
        )
    else:
        st.success("✅ No missing invoices in Vendor.")
    # ====== PAYMENTS ======
    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**💼 ERP Payments**")
        if not erp_pay.empty:
            st.dataframe(
                erp_pay.style.applymap(lambda _: "background-color: #004d40; color: white"),
                use_container_width=True
            )
            st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No ERP payments found.")
        if not erp_pay.empty and not ven_pay.empty:
            total_erp_pay = erp_pay["Amount"].sum()
            total_vendor_pay = ven_pay["Amount"].sum()
            pay_difference = total_erp_pay - total_vendor_pay
            st.markdown("---")
            st.markdown(f"### ⚖️ Difference (ERP - Vendor Payments): **{pay_difference:,.2f} EUR**")
    with col2:
        st.markdown("**🧾 Vendor Payments**")
        if not ven_pay.empty:
            st.dataframe(
                ven_pay.style.applymap(lambda _: "background-color: #1565c0; color: white"),
                use_container_width=True
            )
            st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No Vendor payments found.")
    # ====== MATCHED PAYMENTS ======
    st.subheader("✅ Matched Payments")
    if not matched_pay.empty:
        st.dataframe(
            matched_pay.style.applymap(lambda _: "background-color: #2e7d32; color: white"),
            use_container_width=True
        )
        total_erp = matched_pay["ERP Amount"].sum()
        total_vendor = matched_pay["Vendor Amount"].sum()
        diff_total = abs(total_erp - total_vendor)
        st.markdown(f"**Total Matched ERP Payments:** {total_erp:,.2f} EUR")
        st.markdown(f"**Total Matched Vendor Payments:** {total_vendor:,.2f} EUR")
        st.markdown(f"**Difference Between ERP and Vendor Payments:** {diff_total:,.2f} EUR")
    else:
        st.info("No matching payments found.")
    # ====== UNMATCHED PAYMENTS ======
    st.subheader("⚠️ Unmatched Payments")
    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**💼 Unmatched ERP Payments**")
        unmatched_erp_pay = erp_pay[~erp_pay["reason_erp"].isin(matched_pay["ERP Reason"])] \
            if not erp_pay.empty and not matched_pay.empty else erp_pay
        if not unmatched_erp_pay.empty:
            st.dataframe(
                unmatched_erp_pay.style.applymap(lambda _: "background-color: #f57c00; color: white"),
                use_container_width=True
            )
            st.markdown(f"**Total Unmatched ERP Payments:** {unmatched_erp_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No unmatched ERP payments found.")
    with col4:
        st.markdown("**🧾 Unmatched Vendor Payments**")
        unmatched_ven_pay = ven_pay[~ven_pay["reason_ven"].isin(matched_pay["Vendor Reason"])] \
            if not ven_pay.empty and not matched_pay.empty else ven_pay
        if not unmatched_ven_pay.empty:
            st.dataframe(
                unmatched_ven_pay.style.applymap(lambda _: "background-color: #6a1b9a; color: white"),
                use_container_width=True
            )
            st.markdown(f"**Total Unmatched Vendor Payments:** {unmatched_ven_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No unmatched vendor payments found.")
    
    # ====== DOWNLOAD EXCEL ======
    st.markdown("### 📥 Download Reconciliation Excel Report")
    excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing, matched_pay, tier2_matches)
    st.download_button(
        "💾 Download Excel File",
        data=excel_output,
        file_name="ReconRaptor_Reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
