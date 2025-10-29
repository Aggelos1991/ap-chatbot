import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
from datetime import datetime

# ===== Helper functions =====
def parse_amount(v):
    """Parse numeric strings (EU/US formats) into float."""
    if pd.isna(v):
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
    try:
        return float(s)
    except:
        return 0.0


def find_col(df, names):
    """Find a column name that loosely matches one of the candidates."""
    for c in df.columns:
        name = c.strip().lower().replace(" ", "").replace(".", "")
        for n in names:
            if n.replace(" ", "").replace(".", "").lower() in name:
                return c
    return None


# ===== Streamlit Config =====
st.set_page_config(page_title="The Remitator", layout="wide")
st.title("💀 The Remitator — Hasta la vista, payment remittance. 💀")

# ===== Uploads =====
pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file = st.file_uploader("📂 (Optional) Upload Credit Notes Excel", type=["xlsx"])

# ===== Main Logic =====
if pay_file:
    try:
        df = pd.read_excel(pay_file)
        df.columns = [c.strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Payment file loaded successfully")
    except Exception as e:
        st.error(f"❌ Error loading Payment Excel: {e}")
        st.stop()

    # ---- Required columns in Payment file ----
    req = [
        "Payment Document Code",
        "Alt. Document",
        "Invoice Value",
        "Payment Value",
        "Supplier Name",
        "Supplier's Email",
    ]
    missing = [c for c in req if c not in df.columns]
    if missing:
        st.error(f"❌ Missing columns in Payment Excel: {missing}")
        st.stop()

    pay_code = st.text_input("🔎 Enter Payment Document Code:")
    if not pay_code:
        st.stop()

    subset = df[df["Payment Document Code"].astype(str) == str(pay_code)].copy()
    if subset.empty:
        st.warning("⚠️ No rows found for this Payment Document Code.")
        st.stop()

    # ---- Optional Credit Note file ----
    cn = None
    if cn_file:
        try:
            cn = pd.read_excel(cn_file)
            cn.columns = [c.strip() for c in cn.columns]
            cn = cn.loc[:, ~cn.columns.duplicated()]
            st.info("📄 Credit Note file loaded and will be applied.")
        except Exception as e:
            st.warning(f"⚠️ Error loading CN file (will skip CN logic): {e}")
            cn = None
    else:
        st.info("ℹ️ No Credit Note file uploaded — showing payments only.")

    # ---- Parse numeric columns ----
    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor = subset["Supplier Name"].iloc[0]
    email = subset["Supplier's Email"].iloc[0]

    # ---- Base summary ----
    summary = subset[["Alt. Document", "Invoice Value"]].copy()
    cn_rows = []

    # ---- Apply CN logic only if CN file exists ----
if cn is not None:
        cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document"])
        cn_val_col = find_col(cn, ["Amount"])

    if cn_alt_col and cn_val_col:
        cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)

        # 🔹 Keep only the last occurrence of each CN number
        cn = cn.drop_duplicates(subset=[cn_alt_col], keep="last").reset_index(drop=True)

        used_cns = set()  # Track CNs already included

        for _, row in subset.iterrows():
            payment_val = row["Payment Value"]
            invoice_val = row["Invoice Value"]
            diff = round(payment_val - invoice_val, 2)

            if abs(diff) > 0.01:
                # Find CNs matching the difference, excluding already used ones
                match = cn[
                    (cn[cn[cn_val_col].abs().round(2) == abs(diff)]) &
                    (~cn[cn_alt_col].isin(used_cns))
                ]

                if not match.empty:
                    for _, cn_row in match.iterrows():
                        cn_no = str(cn_row[cn_alt_col])
                        if cn_no in used_cns:
                            continue  # Skip already used CNs
                        cn_amt = -abs(cn_row[cn_val_col])
                        cn_rows.append(
                            {"Alt. Document": f"{cn_no} (CN)", "Invoice Value": cn_amt}
                        )
                        used_cns.add(cn_no)  # Mark as used

    else:
        st.warning("⚠️ CN file missing expected columns ('Alt.Document', 'Amount'). CN logic skipped.")




    # ---- Add CNs ----
    if cn_rows:
        summary = pd.concat([summary, pd.DataFrame(cn_rows)], ignore_index=True)

    # ---- Add total ----
    total_val = summary["Invoice Value"].sum()
    total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_val}])
    summary = pd.concat([summary, total_row], ignore_index=True)

    # ---- Format ----
    summary["Invoice Value (€)"] = summary["Invoice Value"].apply(
        lambda v: f"€{v:,.2f}"
    )
    summary = summary[["Alt. Document", "Invoice Value (€)"]]

    # ---- Display ----
    st.divider()
    st.subheader(f"📋 Summary for Payment Code: {pay_code}")
    st.write(f"**Vendor:** {vendor}")
    st.write(f"**Vendor Email:** {email}")
    st.dataframe(summary, use_container_width=True)

    # ---- Export Excel ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    for r in dataframe_to_rows(summary, index=False, header=True):
        ws.append(r)

    # ---- Hidden meta table ----
    ws_hidden = wb.create_sheet("HiddenMeta")
    meta_data = [
        ["Vendor", vendor],
        ["Vendor Email", email],
        ["Payment Code", pay_code],
        ["Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ]
    for row in meta_data:
        ws_hidden.append(row)

    tab = Table(displayName="MetaTable", ref=f"A1:B{len(meta_data)}")
    style = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    tab.tableStyleInfo = style
    ws_hidden.add_table(tab)
    ws_hidden.sheet_state = "hidden"

    # ---- Save ----
    folder = os.path.join(os.getcwd(), "exports")
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, f"{vendor}_Payment_{pay_code}.xlsx")
    wb.save(file_path)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.download_button(
        "💾 Download Excel Summary",
        buffer,
        file_name=f"{vendor}_Payment_{pay_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("📂 Please upload the Payment Excel to begin (Credit Note file optional).")
