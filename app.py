import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime

# ============== helpers ==============
def parse_amount_to_cents(v):
    """Robust EU/US number parser -> int cents."""
    if pd.isna(v): return None
    s = str(v).strip()
    s = re.sub(r'[^\d,.\-\+]', '', s)
    # if single comma likely decimal (EU): 1.234,56 -> 1234.56
    if s.count(',') == 1 and (s.count('.') == 0 or s.rfind(',') > s.rfind('.')):
        s = s.replace('.', '').replace(',', '.')
    elif s.count(',') > 0 and s.count('.') > 0:
        # assume comma is thousands
        s = s.replace(',', '')
    try:
        return int(round(float(s) * 100))
    except:
        return None

def find_col(df, candidates):
    """Find first column whose normalized name matches any candidate."""
    norm = {c: re.sub(r'\s+', '', c).lower() for c in df.columns}
    for cand in candidates:
        target = re.sub(r'\s+', '', cand).lower()
        for orig, n in norm.items():
            if n == target:
                return orig
    return None

# ============== app ==============
st.set_page_config(page_title="💼 Vendor Payment Reconciliation — Auto CN by Difference", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Auto CN by Difference")

pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("📂 Upload Credit Notes Excel", type=["xlsx"])

if pay_file and cn_file:
    # ---- load ----
    pay = pd.read_excel(pay_file)
    pay.columns = [c.strip() for c in pay.columns]
    pay = pay.loc[:, ~pay.columns.duplicated()]

    cn = pd.read_excel(cn_file)
    cn.columns = [c.strip() for c in cn.columns]
    cn = cn.loc[:, ~cn.columns.duplicated()]

    # ---- required (payment) ----
    REQUIRED = ["Payment Document Code", "Alt. Document", "Invoice Value", "Supplier Name", "Supplier's Email"]
    miss = [c for c in REQUIRED if c not in pay.columns]
    if miss:
        st.error(f"Missing columns in Payment Excel: {miss}")
        st.stop()

    # ---- inputs ----
    pay_code = st.text_input("🔎 Enter Payment Document Code:")
    if not pay_code:
        st.stop()

    # ---- slice by payment code ----
    subset = pay[pay["Payment Document Code"].astype(str) == str(pay_code)].copy()
    if subset.empty:
        st.warning("No rows for this Payment Document Code.")
        st.stop()

    # ---- amounts to cents ----
    subset["Invoice Value (cents)"] = subset["Invoice Value"].apply(parse_amount_to_cents).fillna(0).astype(int)

    # locate payment total column inside the payment file
    pay_col = find_col(
        pay,
        [
            "Payment Value", "Payment Amount", "Paid Amount", "Amount Paid",
            "Payment", "Paid", "Bank Amount", "Transfer Amount", "Remittance Amount",
            "Payment Total", "Total Payment"
        ]
    )

    if pay_col:
        subset["__pay_cents__"] = subset[pay_col].apply(parse_amount_to_cents)
        total_payment_cents = int(pd.Series(subset["__pay_cents__"].dropna()).sum()) if subset["__pay_cents__"].notna().any() else 0
    else:
        total_payment_cents = 0  # no explicit payment column -> assume 0 (no CN deduction possible)

    # ---- base summary (by Alt. Document) ----
    summary = (
        subset.groupby("Alt. Document", as_index=False)["Invoice Value (cents)"]
        .sum()
        .rename(columns={"Invoice Value (cents)":"Value (cents)"})
    )

    # ---- compute difference: payment - invoices ----
    total_invoices_cents = int(summary["Value (cents)"].sum())
    diff_cents = total_payment_cents - total_invoices_cents  # if negative -> we paid less; CN should match abs(diff)

    # ---- CN table mapping ----
    cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document", "Credit Note", "CreditNote", "Document", "Reference", "Doc No"])
    cn_val_col = find_col(cn, ["Amount", "Invoice Value", "Value", "Credit Amount", "CN Amount", "Importe", "Importe N/C"])

    # ---- try to match CN by absolute value of difference ----
    chosen_cn = None
    if cn_alt_col and cn_val_col and diff_cents != 0:
        cn = cn.copy()
        cn["__cents__"] = cn[cn_val_col].apply(parse_amount_to_cents)
        cn = cn.dropna(subset=["__cents__"])
        target = abs(int(diff_cents))
        matches = cn[cn["__cents__"].abs() == target]
        if not matches.empty:
            last = matches.iloc[-1]  # take LAST if multiple
            chosen_cn = {
                "alt": str(last[cn_alt_col]),
                "cents": -abs(int(last["__cents__"]))  # negative line
            }

    # ---- append CN line (if any) ----
    if chosen_cn:
        summary = pd.concat(
            [summary, pd.DataFrame([{"Alt. Document": f"{chosen_cn['alt']} (CN)", "Value (cents)": chosen_cn["cents"]}])],
            ignore_index=True
        )

    # ---- final total ----
    final_total_cents = int(summary["Value (cents)"].sum())
    summary = pd.concat(
        [summary, pd.DataFrame([{"Alt. Document":"TOTAL", "Value (cents)": final_total_cents}])],
        ignore_index=True
    )

    # ---- display ----
    vendor = str(subset["Supplier Name"].dropna().iloc[0])
    email_to = str(subset["Supplier's Email"].dropna().iloc[0])

    show = summary.copy()
    show["Invoice Value"] = show["Value (cents)"].apply(lambda c: f"€{c/100:,.2f}")
    show = show[["Alt. Document", "Invoice Value"]]

    st.subheader(f"📋 Summary for Payment Code: {pay_code}")
    st.write(f"**Vendor:** {vendor}")
    st.write(f"**Vendor Email (from Excel):** {email_to}")
    st.dataframe(show)

    # ---- export ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    for r in dataframe_to_rows(show, index=False, header=True):
        ws.append(r)

    ws_hidden = wb.create_sheet("HiddenMeta")
    ws_hidden["A1"], ws_hidden["B1"] = "Vendor", vendor
    ws_hidden["A2"], ws_hidden["B2"] = "Vendor Email", email_to
    ws_hidden["A3"], ws_hidden["B3"] = "Payment Code", pay_code
    ws_hidden["A4"], ws_hidden["B4"] = "Invoices Total (cents)", total_invoices_cents
    ws_hidden["A5"], ws_hidden["B5"] = "Payment Total (cents)", total_payment_cents
    ws_hidden["A6"], ws_hidden["B6"] = "Diff (cents) = Pay - Inv", diff_cents
    if chosen_cn:
        ws_hidden["A7"], ws_hidden["B7"] = "Chosen CN", chosen_cn["alt"]
        ws_hidden["A8"], ws_hidden["B8"] = "Chosen CN (cents)", chosen_cn["cents"]
    ws_hidden["A9"], ws_hidden["B9"] = "Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_hidden.sheet_state = "hidden"

    folder = os.path.join(os.getcwd(), "exports")
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{re.sub(r'[^A-Za-z0-9]+','_',vendor)}_Payment_{pay_code}.xlsx")
    wb.save(path)

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button(
        "💾 Download Excel Summary",
        buf,
        file_name=os.path.basename(path),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload both the Payment Excel and the Credit Notes Excel to begin.")
