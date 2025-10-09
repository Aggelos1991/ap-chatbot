import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="AP Email Extractor", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable — Vendor Email Manager")

# ========== FUNCTIONS ==========
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


def detect_invalid(df):
    df = combine_emails(df.copy())
    invalid = df[
        df["combined_emails"].isna()
        | (df["combined_emails"].str.strip() == "")
        | (~df["combined_emails"].str.contains("@", case=False, na=False))
        | (df["combined_emails"].str.match(r"^[;]+$", na=False))
    ]
    vendor_col = next((c for c in df.columns if "vendor" in c or "supp_name" in c), None)
    if vendor_col:
        invalid = invalid[[vendor_col, "combined_emails"]].drop_duplicates(subset=[vendor_col])
    return invalid


# ========== MAIN ==========
uploaded = st.file_uploader("📦 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df = safe_excel_to_df(uploaded)

        # base cleanup
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

        # -------- Prompt 1 -----------
        if prompt and "open amounts emails" in prompt.lower():
            vendor_col = next((c for c in df.columns if "vendor" in c or "supp_name" in c), None)
            df = combine_emails(df)
            if "country" not in df.columns:
                df["country"] = "other"
            df["lang"] = df["country"].str.lower().apply(
                lambda x: "ES" if "spain" in x or x.strip() in ["es", "esp", "españa"] else "EN"
            )
            grouped = (
                df.groupby(["lang", vendor_col])["combined_emails"]
                .apply(lambda x: "; ".join(sorted({e.strip() for e in "; ".join(x).split(";") if e.strip()})))
                .reset_index()
            )
            st.write("🇪🇸 **Spanish Vendors**")
            st.dataframe(grouped[grouped["lang"] == "ES"].drop(columns=["lang"]), use_container_width=True)
            st.write("🇬🇧 **English Vendors**")
            st.dataframe(grouped[grouped["lang"] == "EN"].drop(columns=["lang"]), use_container_width=True)

        # -------- Prompt 2 -----------
        elif prompt and any(k in prompt.lower() for k in ["invalid", "missing", "empty emails"]):
            invalid_df = detect_invalid(df)
            if invalid_df.empty:
                st.success("✅ No missing or invalid emails.")
            else:
                st.warning(f"⚠️ Found {len(invalid_df)} vendors missing or invalid emails.")
                st.dataframe(invalid_df, use_container_width=True)
                st.session_state.invalid_df = invalid_df

        # -------- Prompt 3: multiple additions -----------
        elif prompt and prompt.lower().startswith("add multiple emails"):
            # Example input:
            # add multiple emails:
            # Supplier A: email1; accounting1
            # Supplier B: email2; accounting2
            lines = prompt.splitlines()[1:]  # skip first line
            vendor_col = next((c for c in df.columns if "vendor" in c or "supp_name" in c), None)
            updates = []
            for line in lines:
                m = re.match(r"(.+?):\s*(.+)", line.strip())
                if m:
                    supp = m.group(1).strip()
                    emails = [e.strip() for e in m.group(2).split(";") if e.strip()]
                    if emails:
                        idx = df[vendor_col].astype(str).str.lower() == supp.lower()
                        if idx.any():
                            email_main = emails[0]
                            email_acc = emails[1] if len(emails) > 1 else emails[0]
                            if "email" not in df.columns:
                                df["email"] = ""
                            if "accounting_email" not in df.columns:
                                df["accounting_email"] = ""
                            df.loc[idx, "email"] = email_main
                            df.loc[idx, "accounting_email"] = email_acc
                            updates.append(f"✅ {supp}: {email_main}; {email_acc}")
                        else:
                            updates.append(f"⚠️ Supplier '{supp}' not found.")
            st.session_state.df_session = df
            if updates:
                st.write("\n".join(updates))

        # -------- Prompt 4 -----------
        elif prompt and "all spanish" in prompt.lower() and "english" in prompt.lower():
            df = combine_emails(df)
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
            st.write("🇪🇸 **Spanish emails (copy for Outlook)**")
            st.code(es_emails or "No Spanish emails found", language="text")
            st.write("🇬🇧 **English emails (copy for Outlook)**")
            st.code(en_emails or "No English emails found", language="text")

    except Exception as e:
        st.error(f"❌ Error: {e}")
