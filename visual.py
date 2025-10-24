# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io
import xlsxwriter

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Click bar → Raw data | Export → Fancy Excel (Insert Pivot in Excel)**")

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
        if df.shape[1] < 31:
            st.error("Need A to AE.")
            st.stop()

        # Map columns
        df = df.iloc[:, [0, 1, 4, 6, 29, 30]].copy()
        df.columns = ['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']

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

        # Aggregation
        summary = (
            df.groupby('Vendor_Name')
            .apply(lambda g: pd.Series({
                'Total': g['Open_Amount'].sum(),
                'Overdue': g[g['Overdue']]['Open_Amount'].sum(),
                'Not_Overdue': g[~g['Overdue']]['Open_Amount'].sum()
            }))
            .reset_index()
        )
        top20 = summary.nlargest(20, 'Total')

        # Filters
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        with col2:
            vendor_list = ["Top 20"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_list, key="vendor_select")

        # Base data
        base_df = top20 if selected_vendor == "Top 20" else summary[summary['Vendor_Name'] == selected_vendor]

        # Apply filter
        if status_filter == "Overdue Only":
            base_df['Not_Overdue'] = 0
        elif status_filter == "Not Overdue Only":
            base_df['Overdue'] = 0

        # Melt
        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not_Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0]

        # Title
        title = "Top 20 Vendors" if selected_vendor == "Top 20" else selected_vendor

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
            text='Amount',
            height=max(400, len(plot_df) * 45)
        )
        fig.update_traces(texttemplate='€%{text:,.0f}', textposition='inside')
        fig.update_layout(xaxis_title="Amount (€)", yaxis_title="Vendor", legend_title="Status")

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
                raw_details.to_excel(writer, index=False, sheet_name='Raw_Invoices')
            buffer.seek(0)
            st.download_button(
                "Download This Vendor",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("**Click any bar** to see raw invoice lines.")

        # EXPORT FANCY EXCEL (NO PIVOT — BUT READY FOR PIVOT)
        st.markdown("---")
        st.subheader("Export Fancy Excel (Insert Pivot in Excel)")

        def create_fancy_excel(raw_df, summary_df, filename):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Raw data
                raw_df.to_excel(writer, sheet_name='Raw_Data', index=False, startrow=1, header=False)
                # Summary
                summary_df.to_excel(writer, sheet_name='Summary', index=False, startrow=1, header=False)

                workbook = writer.book

                # Format Raw Data
                ws_raw = writer.sheets['Raw_Data']
                header_fmt = workbook.add_format({
                    'bold': True, 'bg_color': '#1f4e79', 'font_color': 'white', 'border': 1
                })
                for col_num, value in enumerate(raw_df.columns):
                    ws_raw.write(0, col_num, value, header_fmt)
                ws_raw.set_column('C:C', 15, workbook.add_format({'num_format': '€#,##0.00'}))
                ws_raw.set_column('B:B', 12, workbook.add_format({'num_format': 'dd/mm/yyyy'}))
                ws_raw.freeze_panes(1, 0)

                # Format Summary
                ws_sum = writer.sheets['Summary']
                for col_num, value in enumerate(summary_df.columns):
                    ws_sum.write(0, col_num, value, header_fmt)
                for col in ['Total', 'Overdue', 'Not_Overdue']:
                    col_idx = summary_df.columns.get_loc(col)
                    ws_sum.set_column(col_idx, col_idx, 15, workbook.add_format({'num_format': '€#,##0.00'}))
                ws_sum.freeze_panes(1, 0)

            buffer.seek(0)
            return buffer

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            buf = create_fancy_excel(df, summary, "all.xlsx")
            st.download_button("Download All Open", data=buf, file_name="All_Open_Fancy.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b:
            buf = create_fancy_excel(df[df['Overdue']], summary[summary['Overdue'] > 0], "overdue.xlsx")
            st.download_button("Download All Overdue", data=buf, file_name="All_Overdue_Fancy.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_c:
            buf = create_fancy_excel(df[~df['Overdue']], summary[summary['Not_Overdue'] > 0], "not.xlsx")
            st.download_button("Download All Not Overdue", data=buf, file_name="All_Not_Overdue_Fancy.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.info("**Open in Excel → Select Raw_Data → Insert → Pivot Table → Done!**")

    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("Upload your Excel → Click bar → See raw data → Export Fancy Excel")
