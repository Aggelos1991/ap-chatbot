import streamlit as st
import pandas as pd
import re

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
            # Greek
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            # Greek
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            # Greek
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            # Greek
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            # Greek (safe only)
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            # Greek
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

    # Ensure debit/credit exist
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

        # Unified multilingual keywords/patterns
        payment_patterns = [
            r"^πληρωμ",             # Greek "Πληρωμή"
            r"^απόδειξη\s*πληρωμ",  # Greek "Απόδειξη πληρωμής"
            r"^payment",            # English: "Payment"
            r"^bank\s*transfer",    # "Bank Transfer"
            r"^trf",                # "TRF ..."
            r"^remesa",             # Spanish
            r"^pago",               # Spanish
            r"^transferencia",      # Spanish
        ]
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

        # Unified multilingual keywords
        payment_words = [
            "pago", "payment", "transfer", "bank", "saldo", "trf",
            "πληρωμή", "μεταφορά", "τράπεζα", "τραπεζικό έμβασμα"
        ]
        credit_words = [
            "credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση","ακυρωτικό","ακυρωτικό παραστατικό"
        ]
        invoice_words = [
            "factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"
        ]

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

    # ====== SCENARIO 1 & 2: MERGE MULTIPLE AND CREDIT NOTES ======
    merged_rows = []
    for inv, group in erp_use.groupby("invoice_erp", dropna=False):
        if group.empty:
            continue

        # If 3 or more entries → take the last (latest)
        if len(group) >= 3:
            group = group.tail(1)

        # If both INV and CN exist for same number → combine
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

    def extract_digits(v):
        return re.sub(r"\D", "", str(v or "")).lstrip("0")

    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    # Add missing cleaner so we can compute e_code / v_code
    def clean_invoice_code(v):
        """
        Normalize invoice strings for comparison:
        - handle structured formats like 2025-FV-00001-001248-01
        - drop common prefixes
        - remove year snippets (20xx)
        - keep only main numeric part
        """
        if not v:
            return ""
        s = str(v).strip().lower().replace(" ", "")
    
        # 🧩 Handle structured invoice patterns early (before dash removal)
        parts = re.split(r"[-_]", s)
        for p in reversed(parts):
            # numeric block with ≥4 digits, skip if it's a year (2020–2039)
            if re.fullmatch(r"\d{4,}", p) and not re.fullmatch(r"20[0-3]\d", p):
                s = p.lstrip("0")  # trim leading zeros (001248 → 1248)
                break
    
        # 🔧 Drop common prefixes (inv, cn, τιμ, αρ, etc.)
        s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no)\W*", "", s)
        s = re.sub(r"20\d{2}", "", s)
        s = re.sub(r"[^a-z0-9]", "", s)
        s = re.sub(r"^0+", "", s)
        s = re.sub(r"[^\d]", "", s)  # keep only digits for comparison
        return s
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
