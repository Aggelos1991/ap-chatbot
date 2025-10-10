import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime

# ===== Helpers =====
def parse_amount(v):
    """Parse European/US style amounts to float."""
    if pd.isna(v):
        return 0.0
    s = str(v).strip()
    s = re.sub(r'[^\d,.\-]', '', s)
    if s.count(',') == 1 and s.count('.') == 1:
        if s.find(',') > s.find('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif s.count(',') == 1:
        s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

def find_col(df, names):
    for c in df.columns:
        cname = c.strip().replace(" ", "").lower()
        for n in names:
            if n.replace(" ", "").lower() in cname:
                return c
    return None

# ===== Streamlit Config =====
st.set_page_config(page_title="💼 Vendor Payment + CN Matcher", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Auto CN Match by Difference")

# ===== Uploads =====
pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file = st.file_uploader("📂 Upload Credit Notes Excel", type=["xlsx"])

if pay_file and cn_file:
    # --- Load files ---
    pay = pd.read_excel(pay_file)
    pay.columns = [c.strip() for c in pay.columns]
    pay = pay.loc[:, ~pay.columns.duplicated()]

    cn = pd.read_excel(cn_file)
    cn.columns = [c.strip() for c in cn.columns]
    cn = cn.loc[:, ~cn.columns.duplicated()]

    st.success("✅ Both files loaded successfully")

    # --- Required columns in payment file ---
    req_cols = ["Payment Document Code", "Alt. Document", "Invoice Value", "Supplier Name", "Supplier's Email"]
    missing = [c for c in req_cols if c not in pay.columns]
    if missing:
        st.error(f"Missing columns in Payment Excel: {missing}")
        st.stop()

    pay_code = st.text_input("🔎 Enter Payment Document Code:")
    if not pay_code:
        st.stop()

    subset = pay[pay["Payment Document Code"].astype(str) == str(pay_code)].copy()
    if subset.empty:
        st.warning("⚠️ No rows found for this Payment Document Code.")
        st.stop()

    # --- Detect columns ---
    pay_val_col = find_col(pay, ["Payment Value", "Paid Amount", "Amount Paid", "Payment"])
    inv_val_col = "Invoice Value"
    cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document", "Document", "CN Number"])
    cn_val_col = find_col(cn, ["Amount", "Invoice Value", "Value"])

    if not pay_val_col or not cn_alt_col or not cn_val_col:
        st.error("⚠️ Missing required columns (Payment Value or CN columns).")
        st.stop()

    # --- Clean data ---
    subset[pay_val_col] = subset[pay_val_col].apply(parse_amount)
    subset[inv_val_col] = subset[inv_val_col].apply(parse_amount)
    cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)

    vendor = subset["Supplier Name"].iloc[0]
    email = subset["Supplier's Email"].iloc[0]

    # --- Start summary table ---
    summary = subset[["Alt. Document", inv_val_col]].copy()

    cn_lines = []

    # --- Main logic: detect difference per line ---
    for _, row in subset.iterrows():
        payment_val = row[pay_val_col]
        invoice_val = row[inv_val_col]
        diff = round(payment_val - invoice_val, 2)

        if abs(diff) > 0.01:
            # Find CN with matching value
            match = cn[cn[cn_val_col].abs().round(2) == abs(diff)]
            if not match.empty:
                last = match.iloc[-1]
                cn_no = str(last[cn_alt_col])
                cn_amt = -abs(last[cn_val_col])  # Deduct
                cn_lines.append({"Alt. Document": f"{cn_no} (CN)", inv_val_col: cn_amt})

    # --- Add CN lines ---
    if cn_lines:
        cn_df = pd.DataFrame(cn_lines)
        summary = pd.concat([summary, cn_df], ignore_index=True)

    # --- Total line ---
    total_value = summary[inv_val_col].sum()
    total_row = pd.DataFrame([{"Alt. Document": "TOTAL", inv_val_col: total_value}])
    summary = pd.concat([summary, total_row], ignore_index=True)

    # --- Format for display ---
    summary["Invoice Value (€)"] = summary[inv_val_col].apply(lambda v: f"€{v:,.2f}")
    summary = summary[["Alt. Document", "Invoice Value (€)"]]

    st.divider()
    st.subheader(f"📋 Summary for Payment Code: {pay_code}")
    st.write(f"**Vendor:** {vendor}")
    st.write(f"**Vendor Email:** {email}")
    st.dataframe(summary)

    # --- Export Excel ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    for r in dataframe_to_rows(summary, index=False, header=True):
        ws.append(r)

    ws_hidden = wb.create_sheet("HiddenMeta")
    ws_hidden["A1"], ws_hidden["B1"] = "Vendor", vendor
    ws_hidden["A2"], ws_hidden["B2"] = "Vendor Email", email
    ws_hidden["A3"], ws_hidden["B3"] = "Payment Code", pay_code
    ws_hidden["A4"], ws_hidden["B4"] = "Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_hidden.sheet_state = "hidden"

    folder = os.path.join(os.getcwd(), "exports")
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{vendor}_Payment_{pay_code}.xlsx")
    wb.save(path)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.download_button(
        "💾 Download Excel Summary",
        buffer,
        file_name=f"{vendor}_Payment_{pay_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload both the Payment Excel and the Credit Notes Excel to begin.")
