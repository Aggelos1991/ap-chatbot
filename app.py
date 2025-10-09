import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="AP Overdue Email Manager", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable — Overdue Invoice Manager")

# ================= HELPER FUNCTIONS =================
def safe_excel_to_df(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = next((wb[name] for name in wb.sheetnames if wb[name].max_row > 1 and wb[name].max_column > 1), wb.active)
    data = []
    for row in ws.values:
        safe_row = ["" if cell is None else str(cell) for cell in row]
        data.append(safe_row)
    if not data:
        raise ValueError("Excel file is empty.")
    headers = [str(h).strip().lower().replace(" ", "_") if h else f"col_{i}" for i, h in enumerate(data[0])]
    seen = {}
    unique_headers = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            unique_headers.append(h)
    return pd.DataFrame(data[1:], columns=unique_headers)

def combine_emails(df):
    email_cols = [c for c in df.columns if "email" in c]
    if not email_cols:
        return None
    df["combined_emails"] = df[email_cols].apply(
        lambda r: "; ".join(sorted({str(x).strip() for x in r if str(x).strip()})), axis=1
    )
    return df

# ================= MAIN APP =================
uploaded = st.file_uploader("📦 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df = safe_excel_to_df(uploaded)

        # --- FILTER BASE EXCEL ---
        if "document" in df.columns:
            df = df[~df["document"].str.contains("F&B", case=False, na=False)]
        if "type" in df.columns:
            df = df[df["type"].str.upper() == "XPI"]
        for c in [c for c in df.columns if "payment_method" in c]:
            df = df[~df[c].str.lower().isin(["downpayment", "direct debit", "cash", "credit card"])]

        agreed_col = next((c for c in ["agreed", "agreeded"] if c in df.columns), None)
        if agreed_col:
            df[agreed_col] = pd.to_numeric(df[agreed_col], errors="coerce").fillna(0)
            df = df[df[agreed_col] == 0]

        st.session_state.df_session = df
        st.success(f"✅ Excel loaded and filtered: {len(df)} rows")
        st.dataframe(df.head(20), use_container_width=True)

        prompt = st.text_area("Type your request (supports multi-line):")
        df = st.session_state.df_session.copy()

        # ========== SHOW OVERDUE INVOICES ==========
        if prompt.lower().startswith("show overdue invoices"):
            m = re.search(r"as of\s+(\d{4}-\d{2}-\d{2})", prompt)
            ref_date = pd.to_datetime(m.group(1)) if m else datetime.today()

            if "due_date" not in df.columns:
                st.error("⚠️ No due_date column found.")
            else:
                df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")
                overdue_df = df[df["due_date_parsed"] < ref_date].copy()

                vendor_col = next((c for c in ["vendor_name", "supp_name", "supplier", "vendor"] if c in overdue_df.columns), None)
                if "open_amount" in overdue_df.columns:
                    overdue_df["open_amount"] = pd.to_numeric(overdue_df["open_amount"], errors="coerce").fillna(0)
                elif "open_amount_in_base_cur" in overdue_df.columns:
                    overdue_df["open_amount"] = pd.to_numeric(overdue_df["open_amount_in_base_cur"], errors="coerce").fillna(0)
                else:
                    overdue_df["open_amount"] = 0

                st.session_state.filtered_df = overdue_df
                total_overdue = overdue_df["open_amount"].sum()

                st.warning(f"⚠️ Found {len(overdue_df)} overdue invoices as of {ref_date.date()}")
                st.write(f"💰 **Total overdue amount:** {total_overdue:,.2f} EUR")
                display_cols = [c for c in [vendor_col, "document", "due_date", "open_amount"] if c]
                st.dataframe(overdue_df[display_cols], use_container_width=True)

        # ========== GET EMAILS FOR CURRENT FILTER ==========
        elif "get emails for current filter" in prompt.lower():
            if "filtered_df" not in st.session_state or st.session_state.filtered_df.empty:
                st.error("⚠️ No active filter. Run 'show overdue invoices ...' first.")
            else:
                df = combine_emails(st.session_state.filtered_df.copy())
                if "country" not in df.columns:
                    df["country"] = "other"
                df["lang"] = df["country"].str.lower().apply(
                    lambda x: "ES" if "spain" in x or x.strip() in ["es", "esp", "españa"] else "EN"
                )

                es_emails = "; ".join(sorted({
                    e.strip() for e in df.loc[df["lang"] == "ES", "combined_emails"].str.split(";").sum() if e.strip()
                }))
                en_emails = "; ".join(sorted({
                    e.strip() for e in df.loc[df["lang"] == "EN", "combined_emails"].str.split(";").sum() if e.strip()
                }))

                st.write(f"📅 Filtered overdue invoices: {len(df)} rows")
                st.write("🇪🇸 Spanish vendor emails:")
                st.code(es_emails or "No Spanish emails found", language="text")
                st.write("🇬🇧 English vendor emails:")
                st.code(en_emails or "No English emails found", language="text")

        # ========== SHOW TOTAL OPEN AMOUNTS ==========
        elif "show total open amounts" in prompt.lower():
            amount_col = "open_amount" if "open_amount" in df.columns else "open_amount_in_base_cur"
            total = pd.to_numeric(df[amount_col], errors="coerce").sum()
            st.write(f"💰 Total open amount: {total:,.2f} EUR")

        # ========== FIND INVALID OR MISSING EMAILS ==========
        elif "find invalid or missing emails" in prompt.lower():
            email_cols = [c for c in df.columns if "email" in c]
            if not email_cols:
                st.error("⚠️ No email columns found.")
            else:
                invalid_df = df[df[email_cols].apply(lambda x: all(not re.search(r"@.+\.", str(i)) for i in x), axis=1)]
                vendor_col = next((c for c in ["vendor_name", "supp_name", "supplier", "vendor"] if c in df.columns), None)
                cols = [vendor_col] + email_cols if vendor_col else email_cols
                st.dataframe(invalid_df[cols].drop_duplicates(), use_container_width=True)

        # ========== ADD MULTIPLE EMAILS ==========
        elif prompt.lower().startswith("add multiple emails:"):
            lines = prompt.split("\n")[1:]
            vendor_col = next((c for c in ["vendor_name", "supp_name", "supplier", "vendor"] if c in df.columns), None)
            if not vendor_col:
                st.error("⚠️ No vendor column found.")
            else:
                updates = []
                for line in lines:
                    if ":" in line:
                        name, emails = line.split(":", 1)
                        name, emails = name.strip(), emails.strip()
                        mask = df[vendor_col].str.lower().eq(name.lower())
                        df.loc[mask, [c for c in df.columns if "email" in c]] = emails
                        updates.append(name)
                st.session_state.df_session = df
                st.success(f"✅ Updated emails for: {', '.join(updates)}")

    except Exception as e:
        st.error(f"❌ Error: {e}")
