# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Uses PivotTable13 from sheet 'Outstanding Invoices IB' — Currency: €**")

# Session state
if 'selected_vendor' not in st.session_state:
    st.session_state.selected_vendor = None

# Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try = True
    try:
        # Read sheet
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None)

        # Find header "VENDOR"
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)
        if df.shape[1] < 31:
            st.error("Need A to AE.")
            st.stop()

        # Map columns
        df = df.iloc[:, [0, 1, 4, 6, 29, 30]].copy()
        df.columns = ['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']

        # Clean
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_A Reference', errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open invoices.")
            st.stop()

        # Overdue
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # Aggregation
        def agg_vendor(group):
            total = group['Open_Amount'].sum()
            overdue = group[group['Overdue']]['Open_Amount'].sum()
            not_overdue = total - overdue
            return pd.Series({'Total': total, 'Overdue_Amount': overdue, 'Not_Overdue_Amount': not_overdue})
        summary = df.groupby('Vendor_Name').apply(agg_vendor).reset_index()
        top10 = summary.nlargest(10, 'Total')

        # Filters
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        vendor_list = ["Top 10"] + sorted(df['Vendor_Name'].unique().tolist())
        selected_vendor = st.selectbox("Select Vendor", vendor_list, key="vendor_select")

        # Update session state
        if selected_vendor != "Top 10":
            st.session_state.selected_vendor = selected_vendor
        else:
            st.session_state.selected_vendor = None

        # Base data
        base_df = top10 if selected_vendor == "Top 10" else summary[summary['Vendor_Name'] == selected_vendor]

        # Apply filter
        if status_filter == "Overdue Only":
            base_df['Not_Overdue_Amount'] = 0
        elif status_filter == "Not Overdue Only":
            base_df['Overdue_Amount'] = 0

        # Melt
        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue_Amount', 'Not_Overdue_Amount'],
            var_name='Type',
            value_name='Amount'
        ).replace({'Overdue_Amount': 'Overdue', 'Not_Overdue_Amount': 'Not Overdue'})
        plot_df = plot_df[plot_df['Amount'] > 0]

        # Title
        title = "Top 10 Vendors by Open Amount" if selected_vendor == "Top 10" else f"{selected_vendor}"

        # Bar chart with € and click
        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title,
            color_discrete_map={'Overdue': '#FF5252', 'Not Overdue': '#4CAF50'},
            text='Amount',
            height=max(400, len(plot_df) * 50)
        )
        fig.update_traces(texttemplate='€%{text:,.0f}', textposition='inside')
        fig.update_layout(
            xaxis_title="Amount (€)",
            yaxis_title="Vendor",
            legend_title="Status",
            hovermode="y unified"
        )

        # Click handler
        def handle_click(trace, points, state):
            if points.point_inds:
                vendor = plot_df.iloc[points.point_inds[0]]['Vendor_Name']
                st.session_state.selected_vendor = vendor

        st.plotly_chart(fig, use_container_width=True, on_select="rerun", selection_mode="points", key="chart")

        # Show details
        show_vendor = st.session_state.selected_vendor
        if show_vendor:
            st.subheader(f"Details: {show_vendor}")
            details = df[df['Vendor_Name'] == show_vendor].copy()
            details = details[['VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Vendor_Email', 'Account_Email']]
            details['Due_Date'] = details['Due_Date'].dt.strftime('%Y-%m-%d')
            details['Open_Amount'] = details['Open_Amount'].map('€{:,.2f}'.format)
            st.dataframe(details, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                details.to_excel(writer, index=False, sheet_name='Invoices')
            buffer.seek(0)
            st.download_button(
                "Download Details (Excel)",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_open_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("**Click a bar** or **select a vendor** to view details.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("Upload your Excel → App reads **PivotTable13** from **Outstanding Invoices IB**")
