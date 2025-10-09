import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="AP Email Extractor", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable — Vendor Emails by Language")

# ======================
# FILE UPLOAD
# ======================
uploaded = st.file_uploader("📦 Upload Excel (.xlsx)", type=["xlsx"])

def safe_excel_to_df(uploaded_file):
    """Read Excel safely and return a cleaned DataFrame"""
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

    headers = [str(h).strip().lower().replace(" ", "_") if h else f"col_{i}" for i, h in enumerate(data[0])]
    df = pd.DataFrame(data[1:], columns=headers)
    return df

# ======================
# MAIN LOGIC
# ======================
if uploaded:
    try:
        df = safe_excel_to_df(uploaded)

        # --- Clean data ---
        if "document" in df.columns:
            df = df[~df["document"].str.contains("F&B", case=False, na=False)]

        if "type" in df.columns:
            df = df[df["type"].str.upper() == "XPI"]

        if "payment_method_descri" in df.columns:
            df = df[~df["payment_method_descri"].str.lower().isin(
                ["downpayment", "direct debit", "cash", "credit card"]
            )]

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

        # ======================
        # PROMPT
        # ======================
        prompt = st.text_input("Type your request:")
        if prompt and "open amounts emails" in prompt.lower():
            if "vendor_name" not in df.columns or "vendor_email" not in df.columns:
                st.error("⚠️ Missing 'vendor_name' or 'vendor_email' columns.")
            else:
                if "country" not in df.columns:
                    df["country"] = "other"

                df["lang"] = df["country"].str.lower().apply(
                    lambda x: "ES" if "spain" in x or x.strip() in ["es", "esp", "españa"] else "EN"
                )

                grouped = (
                    df.groupby(["lang", "vendor_name"])["vendor_email"]
                    .apply(lambda x: "; ".join(sorted({e.strip() for e in x if e.strip()})))
                    .reset_index()
                )

                es_df = grouped[grouped["lang"] == "ES"].drop(columns=["lang"])
                en_df = grouped[grouped["lang"] == "EN"].drop(columns=["lang"])

                st.write("🇪🇸 **Spanish Vendors (Spain)**")
                st.dataframe(es_df, use_container_width=True)

                st.write("🇬🇧 **English Vendors (Other Countries)**")
                st.dataframe(en_df, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
