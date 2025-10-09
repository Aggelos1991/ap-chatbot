import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="AP Email Extractor", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable — Vendor Emails by Language")

# ========== FUNCTIONS ==========
def safe_excel_to_df(uploaded_file):
    """Read Excel safely, rename duplicate headers, and return cleaned DataFrame."""
    file_bytes = uploaded_file.getvalue()
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

    ws = None
    for name in wb.sheetnames:
        w = wb[name]
        if w.max_row > 1 and w.max_column > 1:
            ws = w
            break
    if ws is None:
        ws = wb.active

    # Convert all cells to strings safely
    data = []
    for row in ws.values:
        safe_row = []
        for cell in row:
            if cell is None:
                safe_row.append("")
            else:
                safe_row.append(str(cell))
        data.append(safe_row)

    if not data:
        raise ValueError("Excel file is empty.")

    headers = [str(h).strip().lower().replace(" ", "_") if h else f"col_{i}" for i, h in enumerate(data[0])]

    # ---- FIX DUPLICATES ----
    seen = {}
    unique_headers = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            unique_headers.append(h)

    df = pd.DataFrame(data[1:], columns=unique_headers)
    return df


# ========== MAIN ==========
uploaded = st.file_uploader("📦 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df = safe_excel_to_df(uploaded)

        # --- CLEAN DATA ---
        if "document" in df.columns:
            df = df[~df["document"].str.contains("F&B", case=False, na=False)]

        if "type" in df.columns:
            df = df[df["type"].str.upper() == "XPI"]

        # detect payment_method columns even if name slightly differs
        pay_cols = [c for c in df.columns if "payment_method" in c]
        for c in pay_cols:
            df = df[~df[c].str.lower().isin(
                ["downpayment", "direct debit", "cash", "credit card"]
            )]

        # agreed/agreeded column
        agreed_col = None
        for col in ["agreed", "agreeded"]:
            if col in df.columns:
                agreed_col = col
                break
        if agreed_col:
            df[agreed_col] = pd.to_numeric(df[agreed_col], errors="coerce").fillna(0)
            df = df[df[agreed_col] == 0]

        st.success(f"✅ Excel loaded and filtered: {len(df)} rows")
        st.dataframe(df.head(20), use_container_width=True)

        # ========== PROMPT ==========
        prompt = st.text_input("Type your request:")

        # --------- Prompt 1: table output ---------
        if prompt and "open amounts emails" in prompt.lower():
            vendor_col = None
            email_col = None

            for c in df.columns:
                if "vendor" in c or "supp_name" in c:
                    vendor_col = c
                if any(k in c for k in ["email", "correo", "διεύθυνση"]):
                    email_col = c

            if not vendor_col or not email_col:
                st.error("⚠️ Missing vendor or email column in Excel.")
            else:
                if "country" not in df.columns:
                    df["country"] = "other"

                df["lang"] = df["country"].str.lower().apply(
                    lambda x: "ES" if "spain" in x or x.strip() in ["es", "esp", "españa"] else "EN"
                )

                grouped = (
                    df.groupby(["lang", vendor_col])[email_col]
                    .apply(lambda x: "; ".join(sorted({e.strip() for e in x if e.strip()})))
                    .reset_index()
                )

                es_df = grouped[grouped["lang"] == "ES"].drop(columns=["lang"])
                en_df = grouped[grouped["lang"] == "EN"].drop(columns=["lang"])

                st.write("🇪🇸 **Spanish Vendors (Spain)**")
                st.dataframe(es_df, use_container_width=True)

                st.write("🇬🇧 **English Vendors (Other Countries)**")
                st.dataframe(en_df, use_container_width=True)

        # --------- Prompt 2: combined list output ---------
        elif prompt and "all spanish" in prompt.lower() and "english" in prompt.lower():
            if "country" not in df.columns:
                df["country"] = "other"

            df["lang"] = df["country"].str.lower().apply(
                lambda x: "ES" if "spain" in x or x.strip() in ["es", "esp", "españa"] else "EN"
            )

            email_col = None
            for c in df.columns:
                if any(k in c for k in ["email", "correo", "διεύθυνση"]):
                    email_col = c
                    break

            if not email_col:
                st.error("⚠️ Email column not found.")
            else:
                es_emails = "; ".join(sorted({e.strip() for e in df.loc[df["lang"] == "ES", email_col] if e.strip()}))
                en_emails = "; ".join(sorted({e.strip() for e in df.loc[df["lang"] == "EN", email_col] if e.strip()}))

                st.write("🇪🇸 **Spanish emails (copy for Outlook)**")
                st.code(es_emails or "No Spanish emails found", language="text")

                st.write("🇬🇧 **English emails (copy for Outlook)**")
                st.code(en_emails or "No English emails found", language="text")

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
