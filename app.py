# ==========================================================
# The Remitator — FINAL FIXED EXPORT (Header + Tables in One Sheet)
# ==========================================================
import os, re, requests
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from dotenv import load_dotenv

st.set_page_config(page_title="The Remitator", layout="wide")
st.title("💀 The Remitator — Hasta la vista, payment remittance. 💀")

# ====== ENV ======
load_dotenv()
GLPI_URL   = os.getenv("GLPI_URL")
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")

# ====== HELPERS ======
def parse_amount(v):
    if pd.isna(v): return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") == 1:
        if s.find(",") > s.find("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif s.count(",") == 1: s = s.replace(",", ".")
    try: return float(s)
    except: return 0.0

def find_col(df, names):
    for c in df.columns:
        name = c.strip().lower().replace(" ", "").replace(".", "")
        for n in names:
            if n.replace(" ", "").replace(".", "").lower() in name:
                return c
    return None

# ====== MAIN ======
pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("📂 (Optional) Upload Credit Notes Excel", type=["xlsx"])

if pay_file:
    df = pd.read_excel(pay_file)
    df.columns = [c.strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    st.success("✅ Payment file loaded successfully")

    pay_input = st.text_input("🔎 Enter one or more Payment Document Codes (comma-separated):", "")
    if not pay_input.strip():
        st.stop()

    selected_codes = [x.strip() for x in pay_input.split(",") if x.strip()]
    if not selected_codes:
        st.stop()

    combined_html = ""
    combined_vendor_names = []
    export_blocks = []

    for pay_code in selected_codes:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)].copy()
        if subset.empty:
            continue

        subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
        subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)
        vendor = subset["Supplier Name"].iloc[0]
        summary = subset[["Alt. Document", "Invoice Value"]].copy()

        cn_rows, unmatched_invoices = [], []
        if cn_file:
            cn = pd.read_excel(cn_file)
            cn.columns = [c.strip() for c in cn.columns]
            cn = cn.loc[:, ~cn.columns.duplicated()]
            cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document"])
            cn_val_col = find_col(cn, ["Amount", "Debit", "Charge", "Cargo", "Invoice Value", "Invoice Value (€)"])
            if cn_alt_col and cn_val_col:
                cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)
                used = set()
                for _, row in subset.iterrows():
                    inv = str(row["Alt. Document"])
                    diff = round(row["Payment Value"] - row["Invoice Value"], 2)
                    match = False
                    for i, r in cn.iterrows():
                        if i in used: continue
                        val = round(abs(r[cn_val_col]), 2)
                        if val == 0: continue
                        if round(val, 2) == round(abs(diff), 2):
                            cn_rows.append({"Alt. Document": f"{r[cn_alt_col]} (CN)", "Invoice Value": -val})
                            used.add(i); match=True; break
                    if not match and abs(diff) > 0.01:
                        unmatched_invoices.append({"Alt. Document": f"{inv} (Adj. Diff)", "Invoice Value": diff})

        valid_cn_df = pd.DataFrame(cn_rows)
        unmatched_df = pd.DataFrame(unmatched_invoices)
        all_rows = pd.concat([summary, valid_cn_df, unmatched_df], ignore_index=True)
        total_val = subset["Payment Value"].sum()
        all_rows.loc[len(all_rows)] = ["TOTAL", total_val]
        all_rows["Invoice Value (€)"] = all_rows["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
        display_df = all_rows[["Alt. Document", "Invoice Value (€)"]]

        html_table = display_df.to_html(index=False, border=0, justify="center", classes="table")
        combined_html += f"<h4>Payment Code: {pay_code} — Vendor: {vendor}</h4>{html_table}<br>"
        combined_vendor_names.append(vendor)

        # save table for export
        block_header = pd.DataFrame([[f"Payment Code: {pay_code}", f"Vendor: {vendor}"]], columns=["", ""])
        export_blocks.append(block_header)
        export_blocks.append(display_df)
        export_blocks.append(pd.DataFrame([["", ""]], columns=["", ""]))  # spacing

    # --- combine all into one Excel sheet ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Remitator Export"

    # add header
    ws.append(["Payment Codes", ", ".join(selected_codes)])
    ws.append(["Vendors", ", ".join(combined_vendor_names)])
    ws.append(["Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append([])

    # stack all tables
    for block in export_blocks:
        for r in dataframe_to_rows(block, index=False, header=True):
            ws.append(r)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.markdown(combined_html, unsafe_allow_html=True)
    st.download_button(
        "💾 Download Full Excel (Header + Tables)",
        buf,
        file_name=f"Remitator_Full_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
