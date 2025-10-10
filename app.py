import pandas as pd
import streamlit as st
from io import BytesIO

# ===== Streamlit config =====
st.set_page_config(page_title="💼 Vendor Payment Reconciliation Exporter", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Excel Export Tool")

uploaded_file = st.file_uploader("📂 Upload Excel (TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Excel loaded successfully")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel: {e}")
        st.stop()

    # --- REQUIRED columns ---
    REQ = [
        "Payment Document Code",
        "Alt. Document",
        "Invoice Value",
        "Supplier Name",
        "Supplier's Email",
    ]
    missing = [c for c in REQ if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Excel: {missing}")
        st.stop()

    # --- Filter by Payment Code ---
    pay_code = st.text_input("🔎 Enter Payment Document Code:")

    if pay_code:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)]
        if subset.empty:
            st.warning("⚠️ No rows found for this Payment Document Code.")
        else:
            # Clean and summarize
            subset = subset.copy()
            subset["Invoice Value"] = pd.to_numeric(subset["Invoice Value"], errors="coerce").fillna(0)
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()

            # Calculate total and append as last row
            total_value = summary["Invoice Value"].sum()
            total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_value}])
            summary = pd.concat([summary, total_row], ignore_index=True)

            # Get vendor details
            vendor = str(subset["Supplier Name"].dropna().iloc[0])
            email_to = str(subset["Supplier's Email"].dropna().iloc[0])

            # Display results
            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email_to}")
            st.dataframe(summary.style.format({"Invoice Value": "€{:,.2f}".format}))

            # --- Save Excel to memory ---
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                summary.to_excel(writer, index=False, sheet_name="Summary")
            buffer.seek(0)

            # --- Download button ---
            filename = f"{vendor}_Payment_{pay_code}.xlsx"
            st.download_button(
                label="💾 Download Excel Summary",
                data=buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ Ready to download the Excel summary file.")
else:
    st.info("Upload your Excel file to begin.")
