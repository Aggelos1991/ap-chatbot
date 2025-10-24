# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Uses PivotTable13 from sheet 'Outstanding Invoices IB'**")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        # Read the correct sheet
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None)

        # Find PivotTable13: look for "VENDOR" in column A
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A. Check PivotTable13.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        # Ensure we have at least AE (column 31)
        if df.shape[1] < 31:
            st.error("Not enough columns. Need A to AE.")
            st.stop()

        # Map columns by position
        df = df.iloc[:, [0, 1, 4, 6, 29, 30]].copy()  # A, B, E, G, AD, AE
        df.columns = ['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']

        # Convert types
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open invoices in PivotTable13.")
            st.stop()

        # Overdue logic
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # Summary by Vendor
        summary = df.groupby('Vendor_Name').agg(
            Total=('Open_Amount', 'sum'),
            Overdue_Amount=('Open_Amount', lambda x: x[df.loc[x.index, 'Overdue']]),
            Not_Overdue_Amount=('Open_Amount', lambda x: x[~df.loc[x.index, 'Overdue']])
        ).reset_index()
        summary['Not_Overdue_Amount'] = summary['Total'] - summary['Overdue_Amount']
        top10 = summary.nlargest(10, 'Total')

        # Filters
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"])
        with col2:
            vendor_list = ["Top 10"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_list)

        # Apply status filter
        plot_df = summary.copy()
        if status_filter == "Overdue Only":
            plot_df['Not_Overdue_Amount'] = 0
        elif status_filter == "Not Overdue Only":
            plot_df['Overdue_Amount'] = 0

        # Apply vendor selection
        if selected_vendor != "Top 10":
            plot_df = plot_df[plot_df['Vendor_Name'] == selected_vendor]
            chart_title = f"{selected_vendor} - Open Items"
        else:
            plot_df = top10
            chart_title = "Top 10 Vendors by Open Amount"

        # Prepare chart data
        plot_df = plot_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue_Amount', 'Not_Overdue_Amount'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df['Type'] = plot_df['Type'].replace({
            'Overdue_Amount': 'Overdue',
            'Not_Overdue_Amount': 'Not Overdue'
        })

        # Bar Chart
        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=chart_title,
            color_discrete_map={'Overdue': '#FF5252', 'Not Overdue': '#4CAF50'},
            text='Amount',
            height=max(400, len(plot_df) * 45)
        )
        fig.update_traces(texttemplate='$%{text:,.0f}', textposition='inside')
        fig.update_layout(
            xaxis_title="Amount ($)",
            yaxis_title="Vendor",
            legend_title="Status",
            hovermode="y unified"
        )

        # Display chart
        chart = st.plotly_chart(fig, use_container_width=True, key="main_chart")
        click_data = st.session_state.get("main_chart", {}).get("clickData")

        # Determine vendor to show
        show_vendor = selected_vendor if selected_vendor != "Top 10" else None
        if click_data and 'points' in click_data:
            show_vendor = click_data['points'][0]['y']

        # Show details
        if show_vendor:
            st.subheader(f"Details: {show_vendor}")
            details = df[df['Vendor_Name'] == show_vendor].copy()
            details = details[['VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Vendor_Email', 'Account_Email']]
            details['Due_Date'] = details['Due_Date'].dt.strftime('%Y-%m-%d')
            details['Open_Amount'] = details['Open_Amount'].map('${:,.2f}'.format)
            st.dataframe(details, use_container_width=True)

            # Export to Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                details.to_excel(writer, index=False, sheet_name='Invoices')
            buffer.seek(0)
            st.download_button(
                label="Download Details (Excel)",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_open_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Click a bar or select a vendor to view invoice details.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("Upload your Excel → App reads **PivotTable13** from **Outstanding Invoices IB**")
