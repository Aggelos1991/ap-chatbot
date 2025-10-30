import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
from datetime import datetime
from itertools import combinations

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

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor = subset["Supplier Name"].iloc[0]
    email = subset["Supplier's Email"].iloc[0]

    summary = subset[["Alt. Document", "Invoice Value"]].copy()
    cn_rows, debug_rows, unmatched_invoices = [], [], []

    # ============================================================== #
    # ✅ ADVANCED CN LOGIC + INCLUDE UNMATCHED DIFFERENCES IN SUMMARY
    # ============================================================== #
    if cn is not None:
        cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document"])
        cn_val_col = find_col(cn, ["Amount", "Debit", "Charge", "Cargo", "DEBE"])

        if cn_alt_col and cn_val_col:
            cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)
            cn = cn[cn[cn_val_col].abs() > 0.01].reset_index(drop=True)
            cn = cn.drop_duplicates(subset=[cn_alt_col], keep="last").reset_index(drop=True)

            used_indices = set()

            for _, row in subset.iterrows():
                inv = str(row["Alt. Document"])
                payment_val = row["Payment Value"]
                invoice_val = row["Invoice Value"]
                diff = round(payment_val - invoice_val, 2)
                matched_cns = []
                if abs(diff) < 0.01:
                    continue

                match_found = False

                # 1️⃣ Try single CN
                for i, r in cn.iterrows():
                    if i in used_indices:
                        continue
                    if round(abs(r[cn_val_col]), 2) == round(abs(diff), 2):
                        cn_no = str(r[cn_alt_col])
                        cn_amt = -abs(r[cn_val_col])
                        cn_rows.append({"Alt. Document": f"{cn_no} (CN)", "Invoice Value": cn_amt})
                        matched_cns.append(cn_no)
                        used_indices.add(i)
                        match_found = True
                        break

                # 2️⃣ Try 2–3 CN combinations
                if not match_found:
                    available = [(i, abs(r[cn_val_col]), r) for i, r in cn.iterrows() if i not in used_indices]
                    for n in [2, 3]:
                        for combo in combinations(available, n):
                            total = round(sum(x[1] for x in combo), 2)
                            if abs(total - abs(diff)) < 0.05:  # ±0.05 tolerance
                                for i, _, r in combo:
                                    cn_no = str(r[cn_alt_col])
                                    cn_amt = -abs(r[cn_val_col])
                                    cn_rows.append({"Alt. Document": f"{cn_no} (CN)", "Invoice Value": cn_amt})
                                    matched_cns.append(cn_no)
                                    used_indices.add(i)
                                match_found = True
                                break
                        if match_found:
                            break

                # If not matched — record difference as its own row
                if not match_found:
                    unmatched_invoices.append({
                        "Alt. Document": f"{inv} (Unmatched Diff)",
                        "Invoice Value": diff
                    })

                debug_rows.append({
                    "Invoice": inv,
                    "Invoice Value": invoice_val,
                    "Payment Value": payment_val,
                    "Difference": diff,
                    "Matched CNs": ", ".join(matched_cns) if matched_cns else "—",
                    "Matched?": "✅" if match_found else "❌"
                })

            # 🧾 Unused CNs
            unmatched_cns = cn.loc[~cn.index.isin(used_indices), [cn_alt_col, cn_val_col]].copy()
            unmatched_cns.rename(columns={cn_alt_col: "CN Number", cn_val_col: "Amount"}, inplace=True)
            unmatched_cns["Amount"] = unmatched_cns["Amount"].apply(lambda v: f"€{v:,.2f}")

            st.success(f"✅ Applied {len(cn_rows)} CNs (single/combo)")
            debug_df = pd.DataFrame(debug_rows)
            if not debug_df.empty:
                st.subheader("🔍 Debug breakdown — invoice vs. CN matching")
                st.dataframe(debug_df, use_container_width=True)
        else:
            st.warning("⚠️ CN file missing expected columns ('Alt.Document', 'Amount/Debit'). CN logic skipped.")
    else:
        st.info("ℹ️ No Credit Note file uploaded — showing payments only.")

    # ==============================================================
    # ✅ Combine all into final summary (invoices + CNs + unmatched)
    # ==============================================================
    all_rows = summary.copy()
    if cn_rows:
        all_rows = pd.concat([all_rows, pd.DataFrame(cn_rows)], ignore_index=True)
    if unmatched_invoices:
        all_rows = pd.concat([all_rows, pd.DataFrame(unmatched_invoices)], ignore_index=True)

    total_val = all_rows["Invoice Value"].sum()
    total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_val}])
    all_rows = pd.concat([all_rows, total_row], ignore_index=True)

    all_rows["Invoice Value (€)"] = all_rows["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    all_rows = all_rows[["Alt. Document", "Invoice Value (€)"]]

    st.divider()
    st.subheader(f"📋 Final Summary for Payment Code: {pay_code}")
    st.write(f"**Vendor:** {vendor}")
    st.write(f"**Vendor Email:** {email}")
    st.dataframe(all_rows, use_container_width=True)

    # ---- Export Excel ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Final Summary"
    for r in dataframe_to_rows(all_rows, index=False, header=True):
        ws.append(r)

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
