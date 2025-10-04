import re
import pandas as pd
import streamlit as st
from datetime import date

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption(
    "Examples: "
    "'vendor name of inv-1003', "
    "'emails for open', "
    "'due smaller than 2025-05-10', "
    "'due between 2024-10-01 and 2025-02-01', "
    "'open amounts by vendor top 5', "
    "'oldest unpaid invoice'"
)

# ---------------------------
# Column normalization
# ---------------------------
SYNONYMS = {
    "invoice_no":   ["invoice no","invoice number","invoice","inv","inv no","inv#","document no","doc no","document"],
    "vendor_name":  ["vendor","vendor name","supplier","supplier name","proveedor"],
    "vendor_email": ["email","vendor email","supplier email","correo","mail"],
    "status":       ["status","state","paid?","open?","payment status"],
    "amount":       ["amount","total","invoice amount","importe","value"],
    "currency":     ["currency","curr","moneda"],
    "due_date":     ["due date","vencimiento","fecha vencimiento"],
    "payment_date": ["payment date","fecha pago","paid date"],
    "po_number":    ["po","po number","purchase order"],
}
STANDARD_COLS = list(SYNONYMS.keys())

def _clean(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    colmap = {}
    for c in df.columns:
        c_clean = _clean(c)
        mapped = None
        for std, alts in SYNONYMS.items():
            for alt in [std] + alts:
                if _clean(alt) == c_clean:
                    mapped = std
                    break
            if mapped:
                break
        colmap[c] = mapped if mapped else c
    df = df.rename(columns=colmap)
    for sc in STANDARD_COLS:
        if sc not in df.columns:
            df[sc] = None
    return df

# ---------------------------
# Invoice IDs (very tolerant)
# ---------------------------
STOPWORDS = {
    "paid","open","pending","invoice","invoices","inv","unpaid","status",
    "email","emails","mail","for","the","can","you","bring","vendor","vendors","supplier","suppliers",
    "amount","currency","due","payment","date","show","list","over","under","below","older","oldest","newest",
    "what","is","tell","give","please","find","greater","than","less","more","sum","total","before","after",
    "since","on","names","totals","by","top","smaller","bigger","between","and"
}

def detect_invoice_ids(text: str):
    """
    Grab tokens that:
      - are alnum with -_/ allowed,
      - contain at least one letter and one digit,
      - and are not a stopword.
    Handles: inv1003, INV-1003, inv_1003/22, etc.
    """
    text = text.lower()
    # tokens of letters/digits/[-_/]
    tokens = re.findall(r"[a-z0-9][a-z0-9\-_/]*", text)
    out, seen = [], set()
    for tok in tokens:
        if tok in STOPWORDS: 
            continue
        if not re.search(r"[a-z]", tok):  # needs a letter
            continue
        if not re.search(r"[0-9]", tok):  # needs a digit
            continue
        # avoid capturing pure dates accidentally: require a letter
        if tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out

def find_best_invoice_match(df: pd.DataFrame, inv: str):
    if "invoice_no" not in df.columns:
        return pd.DataFrame()
    def norm(x): return re.sub(r"[-_/ \t]", "", str(x).strip().lower())
    target = norm(inv)
    tmp = df.copy()
    tmp["__norm_inv__"] = tmp["invoice_no"].astype(str).apply(norm)
    exact = tmp[tmp["__norm_inv__"] == target]
    if not exact.empty:
        return exact.drop(columns="__norm_inv__", errors="ignore")
    like = tmp[tmp["__norm_inv__"].str.contains(target, na=False)]
    return like.drop(columns="__norm_inv__", errors="ignore")

# ---------------------------
# Dates
# ---------------------------
DATE_RX = r"(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})"

def parse_user_date(s: str):
    # try dayfirst and monthfirst
    for dayfirst in (True, False):
        try:
            return pd.to_datetime(s, dayfirst=dayfirst, errors="raise")
        except Exception:
            pass
    return pd.NaT

BEFORE_WORDS = ("before","earlier than","smaller than","less than","until","by","on or before","<","<=")
AFTER_WORDS  = ("after","later than","greater than","bigger than","from","since","on or after",">",">=")

def extract_date_intent(ql: str):
    """
    Returns dict:
      {"mode": "before"/"after"/"on"/"between", "d1": Timestamp, "d2": Timestamp|None}
    or None if no date in query.
    """
    ql = ql.lower()

    # between A and B
    m_between = re.search(rf"between\s+{DATE_RX}\s+and\s+{DATE_RX}", ql)
    if m_between:
        d1 = parse_user_date(m_between.group(1))
        d2 = parse_user_date(m_between.group(2))
        if pd.notna(d1) and pd.notna(d2):
            if d2 < d1: d1, d2 = d2, d1
            return {"mode":"between","d1":d1,"d2":d2}

    # single date
    m = re.search(DATE_RX, ql)
    if not m:
        return None
    d = parse_user_date(m.group(1))
    if pd.isna(d):
        return {"mode":"invalid","d1":None,"d2":None}

    if any(w in ql for w in BEFORE_WORDS): return {"mode":"before","d1":d,"d2":None}
    if any(w in ql for w in AFTER_WORDS):  return {"mode":"after","d1":d,"d2":None}
    if re.search(r"\bon\b", ql):           return {"mode":"on","d1":d,"d2":None}
    # if they mention due + date but no qualifier, interpret as "on"
    if "due" in ql:                         return {"mode":"on","d1":d,"d2":None}
    return {"mode":"on","d1":d,"d2":None}

# ---------------------------
# Helpers
# ---------------------------
def extract_top_n_generic(ql: str):
    m = re.search(r"(?:top|first)\s*(\d+)", ql)
    return int(m.group(1)) if m else None

def wants_oldest(ql: str): return any(w in ql for w in ("oldest","earliest","older"))
def wants_newest(ql: str): return any(w in ql for w in ("newest","latest","newer"))

def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

def unique_nonempty(series: pd.Series):
    vals = (
        series.dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    return sorted(vals, key=str.lower)

def requested_field(ql: str):
    """
    Detect if the user asks for a specific field.
    Returns one of: vendor_name, vendor_email, amount, due_date, status, currency, None
    """
    ql = ql.lower()
    if "vendor name" in ql or re.search(r"\bvendor\b", ql) and "email" not in ql:
        return "vendor_name"
    if "email" in ql or "emails" in ql:
        return "vendor_email"
    if "amount" in ql or "amounts" in ql:
        return "amount"
    if "due date" in ql or re.search(r"\bdue\b", ql) and re.search(DATE_RX, ql) is None:
        return "due_date"
    if "status" in ql:
        return "status"
    if "currency" in ql:
        return "currency"
    return None

# ---------------------------
# Core
# ---------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "Please upload an Excel first.", None

    ql = q.lower()

    # Prepare working view
    working = df.copy()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")
    working["status_norm"] = working["status"].astype(str).str.lower().str.strip()

    # Status filters
    if any(k in ql for k in ("open","unpaid","pending")):
        working = working[working["status_norm"].str.contains("open|unpaid|pending", na=False)]
    if "paid" in ql and not any(k in ql for k in ("unpaid","not paid","open","pending")):
        working = working[working["status_norm"].str.contains("paid", na=False)]

    # Vendor filter: "for sani"
    vm = re.search(r"\bfor\s+([a-z0-9 ._-]+)", ql)
    if vm:
        needle = vm.group(1).strip()
        working = working[working["vendor_name"].astype(str).str.contains(re.escape(needle), case=False, na=False)]

    # Amount filters
    m_over = re.search(r"(over|above|greater than|>=|more than|bigger than)\s*([0-9][0-9,\.]*)", ql)
    if m_over:
        val = float(m_over.group(2).replace(",", ""))
        working = working[working["amount"] >= val]
    m_under = re.search(r"(under|below|less than|<=|smaller than)\s*([0-9][0-9,\.]*)", ql)
    if m_under:
        val2 = float(m_under.group(2).replace(",", ""))
        working = working[working["amount"] <= val2]

    # Due date intent
    di = extract_date_intent(ql)
    if di is not None:
        mode = di["mode"]
        if mode == "invalid":
            return "⚠️ I couldn't understand that date. Please use formats like 2025-05-10 or 10/05/2025.", None
        d1, d2 = di["d1"], di["d2"]
        if mode == "before":
            working = working[working["due_date_parsed"] < d1]
        elif mode == "after":
            working = working[working["due_date_parsed"] > d1]
        elif mode == "on":
            working = working[working["due_date_parsed"].dt.date == d1.date()]
        elif mode == "between":
            working = working[(working["due_date_parsed"] >= d1) & (working["due_date_parsed"] <= d2)]

    # Oldest / newest by due date
    if "due_date_parsed" in working.columns:
        if wants_oldest(ql):
            ws = working.sort_values("due_date_parsed", ascending=True)
            if ws.empty: return "No invoices found with a valid due date.", None
            r = ws.iloc[0]
            return (
                f"📄 Oldest invoice: **{r.get('invoice_no','-')}** — **{r.get('vendor_name','-')}**, "
                f"due **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                f"status **{r.get('status','-')}**.",
                ws.head(1)[["invoice_no","vendor_name","due_date","amount","currency","status"]]
            )
        if wants_newest(ql):
            ws = working.sort_values("due_date_parsed", ascending=False)
            if ws.empty: return "No invoices found with a valid due date.", None
            r = ws.iloc[0]
            return (
                f"📄 Newest invoice: **{r.get('invoice_no','-')}** — **{r.get('vendor_name','-')}**, "
                f"due **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                f"status **{r.get('status','-')}**.",
                ws.head(1)[["invoice_no","vendor_name","due_date","amount","currency","status"]]
            )

    # Specific invoice?
    inv_ids = detect_invoice_ids(ql)
    field = requested_field(ql)

    if inv_ids:
        # Answer about the specific invoice(s)
        rows = []
        texts = []
        for inv in inv_ids:
            found = find_best_invoice_match(working, inv)
            if found.empty:
                texts.append(f"❓ Could not find invoice **{inv}**.")
                continue
            r = found.iloc[0]
            rows.append(r)
            if field == "vendor_name":
                texts.append(f"Vendor name for **{inv}**: **{r.get('vendor_name','-')}**.")
            elif field == "vendor_email":
                texts.append(f"Vendor email for **{inv}**: **{r.get('vendor_email','-')}**.")
            elif field == "amount":
                texts.append(f"Amount for **{inv}**: **{fmt_money(r.get('amount'), r.get('currency'))}**.")
            elif field == "due_date":
                texts.append(f"Due date for **{inv}**: **{r.get('due_date','-')}**.")
            elif field == "status":
                texts.append(f"Status for **{inv}**: **{r.get('status','-')}**.")
            elif field == "currency":
                texts.append(f"Currency for **{inv}**: **{r.get('currency','-')}**.")
            else:
                texts.append(
                    f"Invoice **{r.get('invoice_no','-')}** — vendor **{r.get('vendor_name','-')}**, "
                    f"status **{r.get('status','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                    f"due **{r.get('due_date','-')}**."
                )
        table = pd.DataFrame(rows) if rows else None
        return "\n\n".join(texts), table

    # Field-only outputs on the filtered working set
    if field == "vendor_email":
        emails = unique_nonempty(working["vendor_email"])
        if not emails: return "No vendor emails found for this query.", None
        return f"📧 **{len(emails)} vendor email(s):**\n\n" + "; ".join(emails), pd.DataFrame({"vendor_email":emails})

    if field == "vendor_name":
        names = unique_nonempty(working["vendor_name"])
        if not names: return "No vendor names match your filters.", None
        return f"👤 **{len(names)} vendor(s):**\n\n- " + "\n- ".join(names), pd.DataFrame({"vendor_name":names})

    if field == "amount":
        if working.empty: return "No invoices match your filters.", None
        out = working[["invoice_no","vendor_name","amount","currency"]].copy()
        out["amount"] = out.apply(lambda r: fmt_money(r["amount"], r["currency"]), axis=1)
        return "💵 Amounts for matching invoices:", out[["invoice_no","vendor_name","amount"]]

    if field == "due_date":
        if working.empty: return "No invoices match your filters.", None
        return "🗓️ Due dates for matching invoices:", working[["invoice_no","vendor_name","due_date"]]

    if field == "status":
        if working.empty: return "No invoices match your filters.", None
        return "🏷️ Status for matching invoices:", working[["invoice_no","vendor_name","status"]]

    if field == "currency":
        if working.empty: return "No invoices match your filters.", None
        return "💱 Currencies for matching invoices:", working[["invoice_no","vendor_name","currency"]]

    # Grouped totals by vendor
    if ("vendor" in ql or "supplier" in ql) and any(k in ql for k in ("amounts","amount by","amount per","totals","total by","open amounts")):
        g = (working.groupby("vendor_name", dropna=True)["amount"].sum()
             .reset_index().rename(columns={"amount":"total_amount"})
             .sort_values("total_amount", ascending=False))
        n = extract_top_n_generic(ql)
        if n: g = g.head(n)
        if g.empty: return "No vendor totals for this query.", None
        g["total_amount"] = g["total_amount"].map(lambda x: f"{x:,.2f}")
        return f"📊 Totals by vendor ({'top '+str(n) if n else 'all'}):", g

    # Sum total
    if "sum" in ql or "total" in ql:
        total = pd.to_numeric(working["amount"], errors="coerce").sum()
        return f"💰 Total amount: **{total:,.2f}**", working.reset_index(drop=True)

    if working.empty:
        return "No invoices match your query.", None

    # Default: show filtered set
    return f"Found **{len(working)}** invoice(s) matching your query.", working.reset_index(drop=True)

# ---------------------------
# UI
# ---------------------------
st.sidebar.header("📦 Upload Excel")
st.sidebar.write("Columns: Invoice No, Vendor Name, Vendor Email, Status, Amount, Currency, Due Date, Payment Date, PO Number.")
uploaded = st.file_uploader("Upload your Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        df = pd.read_excel(uploaded, dtype=str)
        df = normalize_columns(df)
        st.session_state.df = df
        st.success("✅ Excel loaded and columns normalized.")
        st.dataframe(df.head(50), use_container_width=True)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")

st.subheader("Chat")
if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

if "history" not in st.session_state:
    st.session_state.history = []

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input(
    "Ask: 'vendor name of inv-1003', 'emails for open', "
    "'due smaller than 2025-05-10', 'due between 2024-10-01 and 2025-02-01', "
    "'open amounts by vendor top 5'"
)
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)

    if isinstance(result_df, pd.DataFrame) and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)
        csv = result_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download results as CSV", csv, file_name="results.csv", mime="text/csv")
