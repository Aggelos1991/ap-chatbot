import streamlit as st
import pandas as pd
import re
import streamlit.components.v1 as components

# ======================================
# BACKGROUND ANIMATION
# ======================================
def set_parallax_bg():
    """Interactive marine parallax background with smooth motion and mouse response."""
    components.html("""
    <canvas id="bgCanvas" style="position:fixed;top:0;left:0;width:100%;height:100%;z-index:-1;"></canvas>
    <script>
    const canvas = document.getElementById('bgCanvas');
    const ctx = canvas.getContext('2d');
    let w, h, gradient;
    let mouseX = 0.5, mouseY = 0.5; // normalized mouse position

    function resize(){
        w = canvas.width = window.innerWidth;
        h = canvas.height = window.innerHeight;
    }
    window.onresize = resize;
    resize();

    document.addEventListener('mousemove', e => {
        mouseX = e.clientX / window.innerWidth;
        mouseY = e.clientY / window.innerHeight;
    });

    function draw(t){
        // gradient shifts based on mouse position
        const offsetX = (mouseX - 0.5) * 200;
        const offsetY = (mouseY - 0.5) * 200;

        gradient = ctx.createLinearGradient(offsetX, offsetY, w + offsetX, h + offsetY);
        gradient.addColorStop(0, '#001F4D');
        gradient.addColorStop(1, '#004f91');
        ctx.fillStyle = gradient;
        ctx.fillRect(0, 0, w, h);

        // moving luminous circles
        ctx.globalAlpha = 0.12;
        ctx.beginPath();
        for(let i = 0; i < 25; i++){
            let x = w/2 + Math.sin(t/1000 + i) * (250 + offsetX/4);
            let y = h/2 + Math.cos(t/1100 + i) * (250 + offsetY/4);
            ctx.arc(x, y, 100, 0, 2 * Math.PI);
        }
        ctx.fillStyle = '#ffffff';
        ctx.fill();
        ctx.globalAlpha = 1;

        requestAnimationFrame(draw);
    }
    requestAnimationFrame(draw);
    </script>
    """, height=0)

# ======================================
# CONFIGURATION
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
set_parallax_bg()

# ======================================
# THEME STYLING
# ======================================
st.markdown("""
<style>
.block-container {
    background-color: rgba(255,255,255,0.65); /* lighter transparency */
    border-radius: 20px;
    box-shadow: 0 8px 30px rgba(0,0,0,0.15);
    padding: 2.5rem 3rem;
    backdrop-filter: blur(10px);
    transition: background 0.4s ease-in-out;
}
[data-testid="stSidebar"] {
    background-color: rgba(255,255,255,0.55);
    backdrop-filter: blur(8px);
}
h1,h2,h3 {
    color:#ffffff;
    font-weight:700;
    text-shadow: 0 2px 8px rgba(0,0,0,0.3);
}
div.stButton>button:first-child {
    background-color:#004f91;
    color:white;
    border:none;
    border-radius:8px;
    padding:0.6rem 1.2rem;
    font-weight:600;
    transition:0.3s ease;
}
div.stButton>button:first-child:hover {
    background-color:#0074cc;
    transform:scale(1.03);
}
</style>
""", unsafe_allow_html=True)
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
# CORE MATCHING LOGIC
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        payment_patterns = [r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer", r"^trf", r"^remesa", r"^pago", r"^transferencia"]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση","ακυρωτικό","ακυρωτικό παραστατικό"]
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
        payment_words = ["pago", "payment", "transfer", "bank", "saldo", "trf", "πληρωμή", "μεταφορά", "τράπεζα", "τραπεζικό έμβασμα"]
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση","ακυρωτικό","ακυρωτικό παραστατικό"]
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

    def extract_digits(v): return re.sub(r"\D", "", str(v or "")).lstrip("0")

    def clean_invoice_code(v):
        if not v: return ""
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
            len_diff = abs(len(e_code) - len(v_code))
            suffix_ok = len(e_code) > 2 and len(v_code) > 2 and len_diff <= 2 and (e_code.endswith(v_code) or v_code.endswith(e_code))
            same_type = (e["__doctype"] == v["__doctype"])
            take_it = same_type and (same_full or (same_clean or suffix_ok) and amt_close)
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

    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]].rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]].rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    return matched_df, missing_in_erp, missing_in_vendor

# ======================================
def extract_payments(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    payment_keywords = ["πληρωμή", "payment", "bank transfer", "transferencia bancaria", "transfer", "trf", "remesa", "pago", "deposit", "μεταφορά", "έμβασμα"]
    exclude_keywords = ["τιμολόγιο", "invoice", "παραστατικό", "έξοδα", "expenses", "expense", "invoice of expenses", "expense invoice", "τιμολόγιο εξόδων", "διόρθωση", "διορθώσεις", "correction", "reclass", "adjustment", "μεταφορά υπολοίπου", "balance transfer"]

    def is_real_payment(reason: str) -> bool:
        text = str(reason or "").lower()
        return any(k in text for k in payment_keywords) and not any(b in text for b in exclude_keywords)

    erp_pay = erp_df[erp_df["reason_erp"].apply(is_real_payment)] if "reason_erp" in erp_df else pd.DataFrame()
    ven_pay = ven_df[ven_df["reason_ven"].apply(is_real_payment)] if "reason_ven" in ven_df else pd.DataFrame()

    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(lambda r: abs(normalize_number(r.get("debit_erp")) - normalize_number(r.get("credit_erp"))), axis=1)
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(lambda r: abs(normalize_number(r.get("debit_ven")) - normalize_number(r.get("credit_ven"))), axis=1)

    matched_payments, used_vendor = [], set()
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

    st.success("✅ Reconciliation complete")

    def highlight_row(row):
        if row["Status"] == "Match": return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference": return ['background-color: #f9a825; color: black'] * len(row)
        return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")
        
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

    st.markdown("### ✅ Matched Payments")
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
