import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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

def normalize_columns(df, tag):
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número", "document", "doc", "ref", "referencia",
            "nº factura", "num factura", "alternative document", "αρ.", "αριθμός", "νουμερο", "νούμερο", "no",
            "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito", "abono", "abonos",
            "importe haber", "valor haber", "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "μonto", "amount", "document value",
            "charge", "total", "totale", "totales", "totals", "base imponible", "importe factura", "importe neto",
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción", "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion", "αιτιολογία", "περιγραφή", "παρατηρήσεις",
            "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού"
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
            r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer",
            r"^trf", r"^remesa", r"^pago", r"^transferencia", r"έμβασμα\s*από\s*πελάτη\s*χειρ."
        ]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]
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
        payment_words = ["pago","payment","transfer","bank","saldo","trf","πληρωμή","μεταφορά","τράπεζα","τραπεζικό έμβασμα"]
        credit_words = ["credit","nota","abono","cn","πιστωτικό","πίστωση","ακυρωτικό"]
        invoice_words = ["factura","invoice","inv","τιμολόγιο","παραστατικό"]
        if any(k in reason for k in payment_words):
            return "IGNORE"
        elif any(k in reason for k in credit_words) or credit > 0:
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

    def clean_invoice_code(v):
        if not v:
            return ""
        s = str(v).strip().lower()
        s = re.sub(r"[^a-z0-9]", "", s)
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
            same_type = (e["__doctype"] == v["__doctype"])
            same_clean = (e_code == v_code)
            if same_type and same_clean:
                matched.append({
                    "ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt, "Vendor Amount": v_amt,
                    "Difference": diff, "Status": "Match" if amt_close else "Difference"
                })
                used_vendor_rows.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = {m["ERP Invoice"] for _, m in matched_df.iterrows()}
    matched_ven = {m["Vendor Invoice"] for _, m in matched_df.iterrows()}
    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]]
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]]
    missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    return matched_df, missing_in_erp, missing_in_vendor

# ======================================
# 🔹 TIER-2 MATCHING
# ======================================
def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""
    s = str(v).strip().replace(".", "/").replace("-", "/")
    try:
        d = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(d):
            return ""
        return d.strftime("%Y-%m-%d")
    except:
        return ""

def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def tier2_match(erp_missing, ven_missing):
    if erp_missing.empty or ven_missing.empty:
        return pd.DataFrame(), ven_missing.copy()
    e_df = erp_missing.rename(columns={"Invoice": "invoice_erp", "Amount": "__amt"}).copy()
    v_df = ven_missing.rename(columns={"Invoice": "invoice_ven", "Amount": "__amt"}).copy()
    e_df["date_norm"] = e_df["Date"].apply(normalize_date) if "Date" in e_df.columns else ""
    v_df["date_norm"] = v_df["Date"].apply(normalize_date) if "Date" in v_df.columns else ""
    matches, used_v = [], set()
    for e_idx, e in e_df.iterrows():
        e_inv, e_amt, e_date = str(e.get("invoice_erp", "")), round(float(e.get("__amt", 0)), 2), e.get("date_norm", "")
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v:
                continue
            v_inv, v_amt, v_date = str(v.get("invoice_ven", "")), round(float(v.get("__amt", 0)), 2), v.get("date_norm", "")
            diff, sim = abs(e_amt - v_amt), fuzzy_ratio(e_inv, v_inv)
            if diff < 0.05 and (e_date == v_date or sim >= 0.8):
                matches.append({
                    "ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt, "Vendor Amount": v_amt,
                    "Difference": diff, "Fuzzy Score": round(sim, 2), "Date": e_date or v_date, "Match Type": "Tier-2"
                })
                used_v.add(v_idx)
                break
    return pd.DataFrame(matches), v_df[~v_df.index.isin(used_v)].copy()

# ======================================
# PAYMENTS
# ======================================
def extract_payments(erp_df, ven_df):
    payment_keywords = ["πληρωμή","payment","bank transfer","transferencia","transfer","trf","remesa","pago","deposit","μεταφορά","έμβασμα"]
    exclude_keywords = ["τιμολόγιο","invoice","παραστατικό","έξοδα","expense","correction","adjustment"]
    
    def is_real_payment(r):
        t = str(r or "").lower()
        return any(k in t for k in payment_keywords) and not any(b in t for b in exclude_keywords)
    
    erp_pay = erp_df[erp_df.get("reason_erp","").apply(is_real_payment)] if "reason_erp" in erp_df else pd.DataFrame()
    ven_pay = ven_df[ven_df.get("reason_ven","").apply(is_real_payment)] if "reason_ven" in ven_df else pd.DataFrame()
    
    for d, col in [(erp_pay,"erp"),(ven_pay,"ven")]:
        if not d.empty:
            d["Amount"]=d.apply(lambda r:abs(normalize_number(r.get(f"debit_{col}"))-normalize_number(r.get(f"credit_{col}"))),axis=1)
    
    matched=[]
    used=set()
    for _,e in erp_pay.iterrows():
        for vi,v in ven_pay.iterrows():
            if vi in used:
                continue
            diff=abs(e["Amount"]-v["Amount"])
            if diff<0.05:
                matched.append({
                    "ERP Reason":e.get("reason_erp",""),
                    "Vendor Reason":v.get("reason_ven",""),
                    "ERP Amount":e["Amount"],
                    "Vendor Amount":v["Amount"],
                    "Difference":diff
                })
                used.add(vi)
                break
    return erp_pay,ven_pay,pd.DataFrame(matched)

# ======================================
# ELEGANT EXCEL EXPORT
# ======================================
def style_header(ws, start_row, end_col, header_color, font_color="white"):
    """Style header row with color and bold font"""
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_font = Font(bold=True, color=font_color, size=11)
    header_align = Alignment(horizontal="center", vertical="center")
    
    for col in range(1, end_col + 1):
        cell = ws.cell(row=start_row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        # Add thin border
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        cell.border = thin_border

def style_data_row(ws, start_row, end_row, end_col, row_color):
    """Style data rows with alternating colors"""
    data_fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")
    data_font = Font(size=10)
    
    for row in range(start_row, end_row + 1):
        for col in range(1, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = data_fill
            cell.font = data_font
            # Add thin border
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            cell.border = thin_border

def add_summary_stats(ws, start_row, col):
    """Add summary statistics"""
    summary_font = Font(bold=True, size=12, color="000000")
    ws.cell(row=start_row, column=col, value="SUMMARY STATISTICS").font = summary_font
    return start_row + 2

def export_reconciliation_excel(matched, erp_missing, ven_missing, tier2_matches, erp_pay, ven_pay, matched_pay):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1. MATCHED & DIFFERENCES SHEET
        if not matched.empty:
            matched.to_excel(writer, index=False, sheet_name="Matched & Differences")
            ws1 = writer.sheets["Matched & Differences"]
            style_header(ws1, 1, len(matched.columns), "#2e7d32")  # Green header
            
            # Style Match rows (Green)
            match_mask = matched["Status"] == "Match"
            match_rows = matched[match_mask].index.tolist()
            for row_idx in match_rows:
                style_data_row(ws1, row_idx + 2, row_idx + 2, len(matched.columns), "#e8f5e8")
            
            # Style Difference rows (Orange)
            diff_mask = matched["Status"] == "Difference"
            diff_rows = matched[diff_mask].index.tolist()
            for row_idx in diff_rows:
                style_data_row(ws1, row_idx + 2, row_idx + 2, len(matched.columns), "#fff3e0")
        
        # 2. TIER-2 MATCHES SHEET
        if not tier2_matches.empty:
            tier2_matches.to_excel(writer, index=False, sheet_name="Tier-2 Matches")
            ws2 = writer.sheets["Tier-2 Matches"]
            style_header(ws2, 1, len(tier2_matches.columns), "#2196f3")  # Blue header
            style_data_row(ws2, 2, len(tier2_matches), len(tier2_matches.columns), "#e3f2fd")
        
        # 3. MISSING INVOICES SHEET
        ws3 = writer.book.create_sheet("Missing Invoices")
        start_row = 1
        
        # ERP Missing Header
        ws3.cell(row=start_row, column=1, value="MISSING IN ERP").font = Font(bold=True, size=14, color="#c62828")
        start_row += 2
        erp_missing.to_excel(writer, index=False, sheet_name="Missing Invoices", 
                           startrow=start_row-1, startcol=1)
        style_header(ws3, start_row, len(erp_missing.columns), "#c62828")  # Red header
        style_data_row(ws3, start_row+1, start_row + len(erp_missing), len(erp_missing.columns), "#ffebee")
        start_row += len(erp_missing) + 3
        
        # Vendor Missing Header
        ws3.cell(row=start_row, column=1, value="MISSING IN VENDOR").font = Font(bold=True, size=14, color="#c62828")
        start_row += 2
        ven_missing.to_excel(writer, index=False, sheet_name="Missing Invoices", 
                           startrow=start_row-1, startcol=1)
        style_header(ws3, start_row, len(ven_missing.columns), "#c62828")  # Red header
        style_data_row(ws3, start_row+1, start_row + len(ven_missing), len(ven_missing.columns), "#ffebee")
        
        # 4. PAYMENTS SHEET
        ws4 = writer.book.create_sheet("Payments")
        start_row = 1
        
        # ERP Payments
        ws4.cell(row=start_row, column=1, value="ERP PAYMENTS").font = Font(bold=True, size=14, color="#004d40")
        start_row += 2
        if not erp_pay.empty:
            erp_pay.to_excel(writer, index=False, sheet_name="Payments", startrow=start_row-1, startcol=1)
            style_header(ws4, start_row, len(erp_pay.columns), "#004d40")  # Teal header
            style_data_row(ws4, start_row+1, start_row + len(erp_pay), len(erp_pay.columns), "#e0f2f1")
            start_row += len(erp_pay) + 3
        
        # Vendor Payments
        ws4.cell(row=start_row, column=1, value="VENDOR PAYMENTS").font = Font(bold=True, size=14, color="#1565c0")
        start_row += 2
        if not ven_pay.empty:
            ven_pay.to_excel(writer, index=False, sheet_name="Payments", startrow=start_row-1, startcol=1)
            style_header(ws4, start_row, len(ven_pay.columns), "#1565c0")  # Blue header
            style_data_row(ws4, start_row+1, start_row + len(ven_pay), len(ven_pay.columns), "#e3f2fd")
            start_row += len(ven_pay) + 3
        
        # Payment Matches
        ws4.cell(row=start_row, column=1, value="PAYMENT MATCHES").font = Font(bold=True, size=14, color="#2e7d32")
        start_row += 2
        if not matched_pay.empty:
            matched_pay.to_excel(writer, index=False, sheet_name="Payments", startrow=start_row-1, startcol=1)
            style_header(ws4, start_row, len(matched_pay.columns), "#2e7d32")  # Green header
            style_data_row(ws4, start_row+1, start_row + len(matched_pay), len(matched_pay.columns), "#e8f5e8")
        
        # 5. SUMMARY SHEET
        summary_data = {
            "Metric": [
                "Total Matched Invoices", "Tier-2 Matches", "Missing in ERP", "Missing in Vendor",
                "Total ERP Payments", "Total Vendor Payments", "Payment Difference"
            ],
            "Value": [
                len(matched), len(tier2_matches), len(erp_missing), len(ven_missing),
                erp_pay["Amount"].sum() if not erp_pay.empty else 0,
                ven_pay["Amount"].sum() if not ven_pay.empty else 0,
                abs(erp_pay["Amount"].sum() - ven_pay["Amount"].sum()) if not erp_pay.empty and not ven_pay.empty else 0
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        ws5 = writer.sheets["Summary"]
        
        # Summary styling
        style_header(ws5, 1, 2, "#424242", "#ffffff")  # Dark gray header
        style_data_row(ws5, 2, len(summary_df), 2, "#f5f5f5")
        
        # Format currency columns
        for row in range(2, len(summary_df) + 2):
            ws5.cell(row=row, column=2).number_format = '#,##0.00 "EUR"'
    
    output.seek(0)
    return output

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
        
        # 🧩 Tier-2 Matching
        tier2_matches, ven_missing_after_tier2 = tier2_match(erp_missing, ven_missing)
        
        # Remove Tier-2 matched items from missing tables
        if not tier2_matches.empty:
            matched_vendor_invoices = tier2_matches["Vendor Invoice"].unique().tolist()
            matched_erp_invoices = tier2_matches["ERP Invoice"].unique().tolist()
            erp_missing = erp_missing[~erp_missing["Invoice"].isin(matched_erp_invoices)]
            ven_missing = ven_missing_after_tier2
        
        st.success("✅ Reconciliation complete")

    # Highlight function for matched table
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color:#2e7d32;color:white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color:#f9a825;color:black'] * len(row)
        return [''] * len(row)

    # Display Matched Invoices
    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")

    # Display Missing in ERP
    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color:#c62828;color:white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in ERP.")

    # Display Missing in Vendor
    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color:#c62828;color:white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in Vendor.")

    # Display Tier-2 section with coloring
    st.markdown("### 🧩 Tier-2 Matching (same date, same value, fuzzy invoice)")
    def highlight_tier2(row):
        return ['background-color:#2196f3;color:white'] * len(row)
    
    if not tier2_matches.empty:
        st.success(f"✅ Tier-2 matched {len(tier2_matches)} additional pairs.")
        st.dataframe(tier2_matches.style.apply(highlight_tier2, axis=1), use_container_width=True)
    else:
        st.info("No Tier-2 matches found.")

    # Payments Section
    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**💼 ERP Payments**")
        if not erp_pay.empty:
            st.dataframe(erp_pay.style.applymap(lambda _: "background-color:#004d40;color:white"), use_container_width=True)
            st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No ERP payments found.")
    
    with col2:
        st.markdown("**🧾 Vendor Payments**")
        if not ven_pay.empty:
            st.dataframe(ven_pay.style.applymap(lambda _: "background-color:#1565c0;color:white"), use_container_width=True)
            st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No Vendor payments found.")

    st.markdown("### ✅ Payment Matches")
    if not matched_pay.empty:
        st.dataframe(matched_pay.style.applymap(lambda _: "background-color:#2e7d32;color:white"), use_container_width=True)
        total_erp, total_vendor = erp_pay["Amount"].sum(), ven_pay["Amount"].sum()
        diff_total = round(abs(total_erp - total_vendor), 2)
        st.markdown(f"**ERP Payments Total:** {total_erp:,.2f} EUR")
        st.markdown(f"**Vendor Payments Total:** {total_vendor:,.2f} EUR")
        st.markdown(f"**Difference:** {diff_total:,.2f} EUR")
    else:
        st.info("No payments found.")

    # Excel Export - NOW SUPER ELEGANT!
    st.markdown("### 📥 Download Reconciliation Excel Report")
    excel_output = export_reconciliation_excel(
        matched, erp_missing, ven_missing, tier2_matches, erp_pay, ven_pay, matched_pay
    )
    st.download_button(
        "⬇️ Download Excel Report",
        data=excel_output,
        file_name="Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
