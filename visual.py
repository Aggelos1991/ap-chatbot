# ===============================================================
# Overdue Invoices – Priority Vendors Dashboard (CLICK ENABLED)
# Includes:
# - AY=0 Filter
# - BFP Mode
# - Email Copy Feature
# - Click Bar → Show Vendor Table Below
# ===============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import io

# === MUST BE FIRST STREAMLIT COMMAND ===
st.set_page_config(page_title="Overdue Invoices", layout="wide")

st.title("Overdue Invoices – Priority Vendors Dashboard")
st.markdown("**Click a bar segment to see vendor details below.**")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()

            keep_cols = [0, 1, 4, 6, 10, 29, 30, 31, 33, 35, 39, 50, 55, 71]  # AY=50, BT=71
            df_raw = pd.read_excel(
                xls, sheet_name='Outstanding Invoices IB', header=None, usecols=keep_cols
            )

        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Alt_Document', 'Vendor_Email', 'Account_Email',
            'AF', 'AH', 'AJ', 'AN', 'AY', 'BD', 'BT'
        ]

        # === FILTERS: YES + BD + AY=0 ===
        yes_mask = (
            (df['AF'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AH'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AJ'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AN'].astype(str).str.strip().str.upper() == 'YES')
        )

        bd_keywords = ['ENTERTAINMENT', 'FALSE', 'REGULAR', 'PRIORITY VENDOR', 'PRIORITY VENDOR OS&E']
        bd_mask = df['BD'].astype(str).str.upper().apply(lambda x: any(k in x for k in bd_keywords))

        def is_zero(v):
            try:
                v = str(v).replace(",", ".").strip()
                return float(v) == 0.0
            except:
                return False

        ay_mask = df['AY'].apply(is_zero)
        df = df[yes_mask & bd_mask & ay_mask].reset_index(drop=True)

        # === CLEAN ===
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No valid invoices after filters.")
            st.stop()

        # === OVERDUE LOGIC ===
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # === MODE SELECTION ===
        mode = st.radio("Select Mode:", ["Priority Vendors", "BFP Only"], horizontal=True)
        if mode == "BFP Only":
            df = df[df['BT'].astype(str).str.upper().str.contains("BFP", na=False)]
            if df.empty:
                st.warning("No BFP invoices found.")
                st.stop()

        # === SUMMARY ===
        summary = (
            df.groupby(['Vendor_Name', 'Status'])['Open_Amount']
            .sum().unstack(fill_value=0).reset_index()
        )
        for c in ['Overdue', 'Not Overdue']:
            if c not in summary.columns:
                summary[c] = 0
        summary['Total'] = summary['Overdue'] + summary['Not Overdue']

        # === FILTERS ===
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show:", ["All Open", "Overdue Only", "Not Overdue Only"])
        with col2:
            top_n_option = st.selectbox("Top N Vendors:", ["Top 20", "Top 30"])

        n = 30 if top_n_option == "Top 30" else 20

        # === TOP N LOGIC ===
        if status_filter == "All Open":
            top_df = summary.nlargest(n, 'Total').copy()
            title = f"{top_n_option} Vendors (All Open)"
        elif status_filter == "Overdue Only":
            top_df = summary.nlargest(n, 'Overdue').copy()
            top_df['Not Overdue'] = 0
            title = f"{top_n_option} Vendors (Overdue Only)"
        else:
            top_df = summary.nlargest(n, 'Not Overdue').copy()
            top_df['Overdue'] = 0
            title = f"{top_n_option} Vendors (Not Overdue Only)"

        # === EMAIL EXPORT ===
        st.markdown("### 📧 Extract Vendor Emails for Outlook")
        vendor_subset = df[df['Vendor_Name'].isin(top_df['Vendor_Name'])].copy()
        emails = pd.concat([
            vendor_subset['Vendor_Email'], vendor_subset['Account_Email']
        ], ignore_index=True).dropna().unique().tolist()
        emails = [e.strip() for e in emails if e.strip() and e.lower() != "nan"]
        email_list = "; ".join(sorted(set(emails)))

        if email_list:
            st.text_area("Emails (ready to copy):", email_list, height=150)
            st.markdown(
                f"""
                <button onclick="navigator.clipboard.writeText('{email_list}')"
                style="background-color:#1f4e79;color:white;border:none;
                padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:bold;">
                📋 Copy to Clipboard</button>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.info("No emails found for this selection.")

        # === PLOT DATA ===
        plot_df = top_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0].copy()

        fig = px.bar(
            plot_df,
            x='Amount', y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title + (" – BFP Only" if mode == "BFP Only" else ""),
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            height=max(500, len(plot_df) * 45),
        )

        fig.update_layout(
            xaxis_title="Amount (€)",
            yaxis_title="Vendor",
            legend_title="Status",
            barmode='stack',
            hovermode='y unified',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=160, r=50, t=80, b=50),
        )

        st.markdown("### 📊 Click a bar to view details below")
        click = st.plotly_chart(fig, use_container_width=True, key="vendor_chart")

        # === HANDLE CLICK EVENTS ===
        click_data = st.session_state.get("vendor_chart.clickData", st.session_state.get("clickData"))
        if click_data is None:
            click_data = st.session_state.get("clickData")

        if click_data:
            try:
                vendor_clicked = click_data["points"][0]["y"]
                st.session_state["clicked_vendor"] = vendor_clicked
            except Exception:
                vendor_clicked = None
        else:
            vendor_clicked = st.session_state.get("clicked_vendor", None)

        st.markdown("---")
        st.subheader("📋 Invoice Details")

        if vendor_clicked:
            vendor_data = df[df['Vendor_Name'] == vendor_clicked].copy()
            vendor_data['Due_Date'] = vendor_data['Due_Date'].dt.strftime('%Y-%m-%d')
            st.dataframe(
                vendor_data[['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Alt_Document']],
                use_container_width=True
            )

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                vendor_data.to_excel(writer, index=False, sheet_name='Vendor_Invoices')
            buf.seek(0)
            st.download_button(
                f"💾 Download {vendor_clicked} Invoices",
                data=buf,
                file_name=f"{vendor_clicked.replace(' ', '_')}_Invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Click any bar above to view its vendor’s invoice details below.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.stop()

else:
    st.info("Upload Excel to start.")
