# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io
import xlsxwriter

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Click bar → Raw data | Export → Filtered Raw Data**")

# Session state
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None

# Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
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
        if df.shape[1] < 56:
            st.error("Need up to BD (column 56).")
            st.stop()

        # Map columns
        df = df.iloc[:, [0, 1, 4, 6, 29, 30, 31, 33, 35, 37, 55]].copy()
        df.columns = ['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email',
                      'AF_Filter', 'AH_Filter', 'AJ_Filter', 'AN_Filter', 'BD_Filter']

        # Clean
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open invoices.")
            st.stop()

        # Overdue
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # === YOUR EXACT FILTERS (ALL ON BY DEFAULT) ===
        st.markdown("---")
        st.subheader("Filters (All ON by default)")

        col1, col2, col3, col4, col5 = st.columns(5)

        # AF = YES only
        with col1:
            af_on = st.checkbox("AF = YES", value=True)
            if af_on:
                df = df[df['AF_Filter'] == 'YES']

        # AH = YES only
        with col2:
            ah_on = st.checkbox("AH = YES", value=True)
            if ah_on:
                df = df[df['AH_Filter'] == 'YES']

        # AJ = YES only
        with col3:
            aj_on = st.checkbox("AJ = YES", value=True)
            if aj_on:
                df = df[df['AJ_Filter'] == 'YES']

        # AN = YES only
        with col4:
            an_on = st.checkbox("AN = YES", value=True)
            if an_on:
                df = df[df['AN_Filter'] == 'YES']

        # BD = ENTERTAINMENT, PRIORITY VENDOR, PRIORITY VENDOR OS&E, REGULAR
        with col5:
            bd_on = st.checkbox("BD = ENTERTAINMENT / PRIORITY / REGULAR", value=True)
            if bd_on:
                allowed_bd = ['ENTERTAINMENT', 'PRIORITY VENDOR', 'PRIORITY VENDOR OS&E', 'REGULAR']
                df = df[df['BD_Filter'].isin(allowed_bd)]

        # === SUMMARY AFTER FILTERS ===
        full_summary = df.groupby(['Vendor_Name', 'Status'])['Open_Amount'].sum().unstack(fill_value=0).reset_index()
        full_summary['Total'] = full_summary['Overdue'] + full_summary['Not Overdue']

        # Filters
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        with col2:
            vendor_list = ["Top 20"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_list, key="vendor_select")

        # GET TOP 20 FOR CURRENT FILTER
        if status_filter == "All Open":
            top_df = full_summary.nlargest(20, 'Total')
            title = "Top 20 Vendors (All Open)"
        elif status_filter == "Overdue Only":
            top_df = full_summary.nlargest(20, 'Overdue')
            top_df['Not Overdue'] = 0
            title = "Top 20 Vendors (Overdue Only)"
        else:
            top_df = full_summary.nlargest(20, 'Not Overdue')
            top_df['Overdue'] = 0
            title = "Top 20 Vendors (Not Overdue Only)"

        # Base data
        base_df = top_df if selected_vendor == "Top 20" else full_summary[full_summary['Vendor_Name'] == selected_vendor]

        # Melt
        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0]

        # Add Total per Vendor
        total_per_vendor = base_df.set_index('Vendor_Name')['Total'].to_dict()
        plot_df['Total'] = plot_df['Vendor_Name'].map(total_per_vendor)

        # Bar chart — MANLY COLORS
        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title,
            color_discrete_map={
                'Overdue': '#8B0000',      # Dark Red
                'Not Overdue': '#4682B4'   # Steel Blue
            },
            height=max(400, len(plot_df) * 50)
        )

        # TOTAL ON TOP
        totals = plot_df.groupby('Vendor_Name')['Amount'].sum().reset_index()
        fig.add_scatter(
            x=totals['Amount'],
            y=totals['Vendor_Name'],
            mode='text',
            text=totals['Amount'].apply(lambda x: f'€{x:,.0f}'),
            textposition='top center',
            textfont=dict(size=14, color='white', family='Arial Black', weight=700),
            showlegend=False,
            hoverinfo='skip'
        )

        fig.update_layout(
            xaxis_title="Amount (€)",
            yaxis_title="Vendor",
            legend_title="Status",
            barmode='stack'
        )

        # Plotly chart
        chart = st.plotly_chart(fig, use_container_width=True, key="vendor_chart", on_select="rerun")

        # Capture click
        if st.session_state.vendor_chart and 'selection' in st.session_state.vendor_chart:
            points = st.session_state.vendor_chart['selection']['points']
            if points:
                st.session_state.clicked_vendor = points[0]['y']

        # Show raw data
        show_vendor = st.session_state.clicked_vendor
        if show_vendor:
            st.subheader(f"Raw Invoices: {show_vendor}")
            raw_details = df[df['Vendor_Name'] == show_vendor].copy()
            raw_details = raw_details[['VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Vendor_Email', 'Account_Email']]
            raw_details['Due_Date'] = raw_details['Due_Date'].dt.strftime('%Y-%m-%d')
            raw_details['Open_Amount'] = raw_details['Open_Amount'].map('€{:,.2f}'.format)
            st.dataframe(raw_details, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                raw_details.to_excel(writer, index=False, sheet_name='Raw_Data')
            buffer.seek(0)
            st.download_button(
                "Download This Vendor",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("**Click any bar** to see raw invoice lines.")

        # EXPORT RAW DATA ONLY
        st.markdown("---")
        st.subheader("Export Filtered Raw Data")

        def export_raw(raw_df, filename):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                raw_df.to_excel(writer, sheet_name='Raw_Data', index=False, startrow=1, header=False)
                workbook = writer.book
                worksheet = writer.sheets['Raw_Data']

                # Header
                header_fmt = workbook.add_format({
                    'bold': True, 'bg_color': '#1f4e79', 'font_color': 'white', 'border': 1
                })
                for col_num, value in enumerate(raw_df.columns):
                    worksheet.write(0, col_num, value, header_fmt)

                # Currency
                worksheet.set_column('C:C', 15, workbook.add_format({'num_format': '€#,##0.00'}))
                worksheet.set_column('B:B', 12, workbook.add_format({'num_format': 'dd/mm/yyyy'}))
                worksheet.freeze_panes(1, 0)

            buffer.seek(0)
            return buffer

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            buf = export_raw(df, "all.xlsx")
            st.download_button("Download All Open", data=buf, file_name="All_Open_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b:
            buf = export_raw(df[df['Overdue']], "overdue.xlsx")
            st.download_button("Download All Overdue", data=buf, file_name="All_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_c:
            buf = export_raw(df[~df['Overdue']], "not.xlsx")
            st.download_button("Download All Not Overdue", data=buf, file_name="All_Not_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("Upload your Excel → Filters → Click bar → Export Raw")
