import streamlit as st
import pandas as pd
import re
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
            "debit", "debe", "cargo", "importe", "importe total", "valor", "μonto",
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
            r"^πληρωμ",
            r"^απόδειξη\s*πληρωμ",
            r"^payment",
            r"^bank\s*transfer",
            r"^trf",
            r"^remesa",
            r"^pago",
            r"^transferencia",
            r"έμβασμα\s*από\s*πελάτη\s*χειρ.",
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

        payment_words = ["pago", "payment", "transfer", "bank", "saldo", "trf", "πληρωμή", "μεταφορά", "τράπεζα", "τραπεζικό έμβασμα","έμβασμα από πελάτη χειρ."]
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό", "ακυρωτικό παραστατικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]

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
            same_full = (e_inv == v_inv)
            same_clean = (e_code == v_code)
            if same_type and (same_full or same_clean):
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
    if "Date" in e_df.columns:
        e_df["date_norm"] = e_df["Date"].apply(normalize_date)
    else:
        e_df["date_norm"] = ""
    if "Date" in v_df.columns:
        v_df["date_norm"] = v_df["Date"].apply(normalize_date)
    else:
        v_df["date_norm"] = ""
    matches = []
    used_v = set()
    for e_idx, e in e_df.iterrows():
        e_inv, e_amt, e_date = str(e.get("invoice_erp", "")), round(float(e.get("__amt", 0)), 2), e.get("date_norm", "")
        for v_idx, v in v_df.iterrows():
            if v_idx in used_v:
                continue
            v_inv, v_amt, v_date = str(v.get("invoice_ven", "")), round(float(v.get("__amt", 0)), 2), v.get("date_norm", "")
            diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_inv, v_inv)
            if diff < 0.05 and (e_date == v_date or sim >= 0.8):
                matches.append({
                    "ERP Invoice": e_inv, "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt, "Vendor Amount": v_amt,
                    "Difference": diff, "Fuzzy Score": round(sim, 2),
                    "Date": e_date or v_date, "Match Type": "Tier-2"
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
        t=str(r or "").lower()
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
            if vi in used: continue
            diff=abs(e["Amount"]-v["Amount"])
            if diff<0.05:
                matched.append({"ERP Reason":e.get("reason_erp",""),"Vendor Reason":v.get("reason_ven",""),
                                "ERP Amount":e["Amount"],"Vendor Amount":v["Amount"],"Difference":diff})
                used.add(vi); break
    return erp_pay,ven_pay,pd.DataFrame(matched)

# ======================================
# STREAMLIT UI
# ======================================
uploaded_erp=st.file_uploader("📂 Upload ERP Export (Excel)",type=["xlsx"])
uploaded_vendor=st.file_uploader("📂 Upload Vendor Statement (Excel)",type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_raw=pd.read_excel(uploaded_erp,dtype=str)
    ven_raw=pd.read_excel(uploaded_vendor,dtype=str)
    erp_df=normalize_columns(erp_raw,"erp")
    ven_df=normalize_columns(ven_raw,"ven")

    with st.spinner("Reconciling invoices..."):
        matched,erp_missing,ven_missing=match_invoices(erp_df,ven_df)
        erp_pay,ven_pay,matched_pay=extract_payments(erp_df,ven_df)

    st.success("✅ Reconciliation complete")

    def highlight_row(row):
        if row["Status"]=="Match": return ['background-color:#2e7d32;color:white']*len(row)
        elif row["Status"]=="Difference": return ['background-color:#f9a825;color:black']*len(row)
        return ['']*len(row)

    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row,axis=1),use_container_width=True)
    else:
        st.info("No matches found.")

    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color:#c62828;color:white"),use_container_width=True)
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color:#c62828;color:white"),use_container_width=True)
    else:
        st.success("✅ No missing invoices in Vendor.")

    # 🧩 Tier-2 Matching Layer
    st.markdown("### 🧩 Tier-2 Matching (same date, same value, fuzzy invoice)")
    if not erp_missing.empty and not ven_missing.empty:
        with st.spinner("Running Tier-2 fuzzy matching..."):
            tier2_matches, still_unmatched = tier2_match(erp_missing, ven_missing)

        if not tier2_matches.empty:
          
            # 🧹 Remove Tier-2 matched invoices from missing lists
            matched_vendor_invoices = tier2_matches["Vendor Invoice"].unique().tolist()
            matched_erp_invoices = tier2_matches["ERP Invoice"].unique().tolist()
            ven_missing = ven_missing[~ven_missing["Invoice"].isin(matched_vendor_invoices)]
            erp_missing = erp_missing[~erp_missing["Invoice"].isin(matched_erp_invoices)]
        
            st.success(f"✅ Tier-2 matched {len(tier2_matches)} additional pairs.")
            st.dataframe(tier2_matches,use_container_width=True)

        else:
            st.info("No Tier-2 matches found.")
    else:
        st.info("Tier-2 matching not applicable — one side has no missing items.")

    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1,col2=st.columns(2)
    with col1:
        st.markdown("**💼 ERP Payments**")
        if not erp_pay.empty:
            st.dataframe(erp_pay.style.applymap(lambda _: "background-color:#004d40;color:white"),use_container_width=True)
            st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No ERP payments found.")
    with col2:
        st.markdown("**🧾 Vendor Payments**")
        if not ven_pay.empty:
            st.dataframe(ven_pay.style.applymap(lambda _: "background-color:#1565c0;color:white"),use_container_width=True)
            st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No Vendor payments found.")

    st.markdown("### ✅ Payment Matches")
    if not matched_pay.empty:
        st.dataframe(matched_pay.style.applymap(lambda _: "background-color:#2e7d32;color:white"),use_container_width=True)
        total_erp, total_vendor = erp_pay["Amount"].sum(), ven_pay["Amount"].sum()
        diff_total = round(abs(total_erp - total_vendor), 2)
        st.markdown(f"**ERP Payments Total:** {total_erp:,.2f} EUR")
        st.markdown(f"**Vendor Payments Total:** {total_vendor:,.2f} EUR")
        st.markdown(f"**Difference:** {diff_total:,.2f} EUR")
    else:
        st.info("No payments found.")

    # Excel Export unchanged
    def export_reconciliation_excel(matched, erp_missing, ven_missing):
        import io
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
        output=io.BytesIO()
        with pd.ExcelWriter(output,engine="openpyxl") as writer:
            matched.to_excel(writer,index=False,sheet_name="Matched & Differences")
            ws1=writer.sheets["Matched & Differences"]
            header_fill=PatternFill(start_color="4CAF50",end_color="4CAF50",fill_type="solid")
            header_font=Font(bold=True,color="FFFFFF")
            for cell in ws1[1]:
                cell.fill=header_fill
                cell.font=header_font
                cell.alignment=Alignment(horizontal="center",vertical="center")
            for i,row in enumerate(ws1.iter_rows(min_row=2),start=2):
                if i%2==0:
                    for cell in row:
                        cell.fill=PatternFill(start_color="E8F5E9",end_color="E8F5E9",fill_type="solid")
            for col in ws1.columns:
                max_len=max(len(str(c.value)) if c.value else 0 for c in col)
                ws1.column_dimensions[get_column_letter(col[0].column)].width=max_len+2
            ws_name="Missing"
            erp_missing.to_excel(writer,index=False,sheet_name=ws_name,startrow=4)
            start_col=len(erp_missing.columns)+4
            ven_missing.to_excel(writer,index=False,sheet_name=ws_name,startcol=start_col,startrow=4)
            ws2=writer.sheets[ws_name]
            ws2["A1"]="Missing in ERP"
            ws2["A1"].font=Font(bold=True,size=14,color="FFFFFF")
            ws2["A1"].fill=PatternFill(start_color="E53935",end_color="E53935",fill_type="solid")
            ws2.cell(row=1,column=start_col+1).value="Missing in Vendor"
            ws2.cell(row=1,column=start_col+1).font=Font(bold=True,size=14,color="FFFFFF")
            ws2.cell(row=1,column=start_col+1).fill=PatternFill(start_color="1E88E5",end_color="1E88E5",fill_type="solid")
            thin=Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))
            for row in ws2.iter_rows(min_row=4):
                for c in row:
                    c.border=thin
            for col in ws2.columns:
                max_len=max(len(str(c.value)) if c.value else 0 for c in col)
                ws2.column_dimensions[get_column_letter(col[0].column)].width=max_len+2
        output.seek(0)
        return output

    st.markdown("### 📥 Download Reconciliation Excel Report")
    excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing)
    st.download_button("⬇️ Download Excel Report",data=excel_output,file_name="Reconciliation_Report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
