# ==========================================================
# The Remitator — OLD FINAL VERSION (Stable Legacy Build)
# ==========================================================
import os, re, requests
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime
from dotenv import load_dotenv

# ========== UI ==========
st.set_page_config(page_title="The Remitator", layout="wide")
st.title("The Remitator — Payment Remittance Generator")

# ========== ENV ==========
load_dotenv()
GLPI_URL = os.getenv("GLPI_URL")
APP_TOKEN = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")

# ===============================
#  PARSE AMOUNTS
# ===============================
def parse_amount(v):
    if pd.isna(v): return 0.0
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
    for c in df.columns:
        name = c.strip().lower().replace(" ", "")
        for n in names:
            if n.replace(" ", "").lower() in name:
                return c
    return None

# ===============================
#  GLPI BASIC FUNCTIONS
# ===============================
def glpi_login():
    r = requests.get(
        f"{GLPI_URL}/initSession",
        headers={"Authorization": f"user_token {USER_TOKEN}", "App-Token": APP_TOKEN}
    )
    return r.json().get("session_token")

def glpi_update_ticket(token, ticket_id, status=5, category_id=None):
    payload = {"input": {"status": status}}
    if category_id:
        payload["input"]["itilcategories_id"] = int(category_id)

    return requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )

def glpi_add_solution(token, ticket_id, body_html):
    payload = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "content": body_html,
            "solutiontypes_id": 10,
            "status": 5
        }
    }
    return requests.post(
        f"{GLPI_URL}/ITILSolution",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )

def glpi_add_followup(token, ticket_id, body_html):
    payload = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "content": body_html,
            "solutiontypes_id": 10
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/ITILFollowup",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )

# ===============================
# MAIN APP
# ===============================
pay_file = st.file_uploader("Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("Upload Credit Notes Excel (optional)", type=["xlsx"])

if not pay_file:
    st.info("Upload Payment Excel to begin.")
    st.stop()

df = pd.read_excel(pay_file)
df.columns = [c.strip() for c in df.columns]
df = df.loc[:, ~df.columns.duplicated()]
st.success("Payment file loaded.")

pay_input = st.text_input("Enter Payment Document Code:", "")
if not pay_input.strip():
    st.stop()

codes = [x.strip() for x in pay_input.split(",") if x.strip()]
if not codes:
    st.stop()

combined_html = ""
vendor_list = []
export_data = {}

for code in codes:
    col = find_col(df, ["Payment Document", "Payment Document Code"])
    if not col:
        st.error("Payment Document column not found")
        st.stop()

    subset = df[df[col].astype(str) == str(code)]
    if subset.empty:
        continue

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor_col = find_col(df, ["Vendor", "Supplier Name", "Supplier"])
    vendor = subset[vendor_col].iloc[0] if vendor_col else "Vendor"

    summary = subset[["Alt. Document", "Invoice Value"]].copy()

    cn_rows = []
    unmatched = []

    if cn_file:
        cn = pd.read_excel(cn_file)
        cn.columns = [c.strip() for c in cn.columns]
        cn = cn.loc[:, ~cn.columns.duplicated()]

        alt_c = find_col(cn, ["Alt Document", "Alt. Document"])
        val_c = find_col(cn, ["Amount", "Invoice Value", "DEBE", "Cargo"])

        if alt_c and val_c:
            cn[val_c] = cn[val_c].apply(parse_amount)
            used = set()

            for _, row in subset.iterrows():
                inv = str(row["Alt. Document"])
                diff = round(row["Payment Value"] - row["Invoice Value"], 2)

                matched = False
                for i, r in cn.iterrows():
                    if i in used: continue
                    if round(abs(r[val_c]),2) == round(abs(diff),2):
                        cn_rows.append({"Alt. Document": f"{r[alt_c]} (CN)", "Invoice Value": -abs(r[val_c])})
                        used.add(i)
                        matched = True
                        break

                if not matched and abs(diff) > 0.01:
                    unmatched.append({"Alt. Document": f"{inv} (Adj. Diff)", "Invoice Value": diff})

    full = pd.concat([
        summary,
        pd.DataFrame(cn_rows),
        pd.DataFrame(unmatched)
    ], ignore_index=True)

    total_val = full["Invoice Value"].sum()
    full.loc[len(full)] = ["TOTAL", total_val]

    export_data[code] = {
        "vendor": vendor,
        "rows": full.copy()
    }

    temp = full.copy()
    temp["Invoice Value (€)"] = temp["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    temp = temp[["Alt. Document", "Invoice Value (€)"]]

    html_table = temp.to_html(index=False, border=0)

    combined_html += f"""
<b>Payment Code:</b> {code}<br>
<b>Vendor:</b> {vendor}<br>
<b>Total Amount:</b> €{total_val:,.2f}<br><br>
{html_table}<br><hr><br>
"""
    vendor_list.append(vendor)

# Clean last HR
combined_html = combined_html.rstrip("<hr><br>")

# ===============================
# SHOW SUMMARY
# ===============================
st.markdown(combined_html, unsafe_allow_html=True)

# ===============================
# EXCEL EXPORT (OLD FINAL)
# ===============================
wb = Workbook()
ws = wb.active
ws.title = "Summary"

ws.append(["The Remitator – Old Final Version"])
ws.append([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
ws.append([f"Payment Codes: {', '.join(codes)}"])
ws.append([f"Vendors: {', '.join(set(vendor_list))}"])
ws.append([])

bold = Font(bold=True)
money_fmt = '#,##0.00 €'

row = 6
for code in codes:
    block = export_data[code]
    vendor = block["vendor"]
    df_block = block["rows"]

    ws.cell(row,1).value = f"Payment Code {code} — {vendor}"
    ws.cell(row,1).font = bold
    row += 2

    ws.cell(row,1).value = "Document"
    ws.cell(row,2).value = "Amount (€)"
    ws.cell(row,1).font = bold
    ws.cell(row,2).font = bold
    row += 1

    for _, r in df_block.iterrows():
        ws.cell(row,1).value = r["Alt. Document"]
        ws.cell(row,2).value = r["Invoice Value"]
        ws.cell(row,2).number_format = money_fmt
        row += 1

    row += 2

# Auto size
for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 2

buf = BytesIO()
wb.save(buf)
buf.seek(0)

st.download_button(
    "Download Excel Summary",
    buf,
    file_name="Remitator_Old_Final.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ===============================
# GLPI (OLD FINAL TEMPLATES)
# ===============================
st.subheader("GLPI Message Sender")

language = st.radio("Language", ["Spanish", "English"], horizontal=True)
ticket_id = st.text_input("Ticket ID")
category_id = st.text_input("Category ID")
assigned_email = st.text_input("Assigned Email (optional)")

if language == "Spanish":
    intro = "Estimado proveedor,<br><br>Adjuntamos las facturas correspondientes a los pagos realizados:<br><br>"
    outro = "<br>Quedamos a su disposición para cualquier aclaración.<br><br>Saludos,<br>Finance"
else:
    intro = "Dear supplier,<br><br>Please find below the invoices corresponding to the executed payments:<br><br>"
    outro = "<br>Should you need any clarification, we remain available.<br><br>Kind regards,<br>Finance Team"

html_message = intro + combined_html + outro
st.markdown(html_message, unsafe_allow_html=True)

if st.button("Send to GLPI"):
    if not ticket_id.isdigit():
        st.error("Invalid Ticket ID")
        st.stop()

    token = glpi_login()
    if not token:
        st.error("GLPI login failed.")
        st.stop()

    glpi_update_ticket(token, ticket_id, status=5, category_id=category_id)
    resp = glpi_add_solution(token, ticket_id, html_message)

    if resp.status_code == 400 or "already solved" in resp.text.lower():
        glpi_add_followup(token, ticket_id, html_message)
        st.warning("Ticket already solved — added as follow-up instead.")
    else:
        st.success("Solution added to GLPI.")
