import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")

# === Session State ===
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None
if 'top_n_option' not in st.session_state:
    st.session_state.top_n_option = "Top 30"

# === File Upload ===
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()

            # Read only needed columns (including BJ = column 61 → index 60)
            keep_cols = [0, 1, 4, 6, 29, 30, 31, 33, 35, 39, 55, 61]
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB',
                                 header=None, usecols=keep_cols)

        # === Find Header Row ===
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header row with 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        # === Assign Column Names ===
        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Vendor_Email', 'Account_Email',
            'AF', 'AH', 'AJ', 'AN', 'BD', 'BJ_Alt_Invoice'
        ]

        # === FILTER 1: All YES in AF, AH, AJ, AN ===
        yes_mask = (
            (df['AF'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AH'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AJ'].astype(str).str.strip().str.upper() == 'YES') &
            (df['AN'].astype(str).str.strip().str.upper() == 'YES')
        )

        # === FILTER 2: BD contains any keyword ===
        bd_keywords = ['ENTERTAINMENT', 'FALSE', 'REGULAR', 'PRIORITY VENDOR', 'PRIORITY VENDOR OS&E']
        bd_mask = df['BD'].astype(str).str.upper().apply(
            lambda x: any(k in x for k in bd_keywords) if pd.notna(x) else False
        )

        # === Apply Filters ===
        df = df[yes_mask & bd_mask].reset_index(drop=True)
        df = df.drop(columns=['AF', 'AH', 'AJ', 'AN', 'BD'])

        if df.empty:
            st.warning("No invoices match the filter criteria (YES in AF/AH/AJ/AN + BD keywords).")
            st.stop()

        # === Clean Data ===
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No valid open invoices after cleaning.")
            st.stop()

        # === Overdue Logic ===
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # === Summary: Total by Vendor & Status ===
        summary = df.groupby(['Vendor_Name', 'Status'])['Open_Amount'].sum().unstack(fill_value=0).reset_index()
        for col in ['Overdue', 'Not Overdue']:
            if col not in summary.columns:
                summary[col] = 0
        summary['Total'] = summary['Overdue'] + summary['Not Overdue']

        # === Filters ===
        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        with col2:
            vendor_options = ["Top N"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_options, key="vendor_select")
        with col3:
            top_n_option = st.selectbox("Show", ["Top 20", "Top 30", "All Vendors"],
                                        index=["Top 20", "Top 30", "All Vendors"].index(st.session_state.top_n_option))
            st.session_state.top_n_option = top_n_option

        # === Top N Logic ===
        top_n = {"Top 20": 20, "Top 30": 30, "All Vendors": len(summary)}.get(top_n_option, 30)
        title_suffix = top_n_option.replace(" Vendors", "")

        # === Apply Status Filter ===
        if status_filter == "All Open":
            plot_data = summary.nlargest(top_n, 'Total').copy()
            title = f"{title_suffix} Vendors (All Open)"
        elif status_filter == "Overdue Only":
            plot_data = summary.nlargest(top_n, 'Overdue').copy()
            plot_data['Not Overdue'] = 0
            title = f"{title_suffix} Vendors (Overdue Only)"
        else:
            if summary['Not Overdue'].sum() == 0:
                st.warning("No 'Not Overdue' invoices.")
                plot_data = summary.head(0).copy()
            else:
                plot_data = summary.nlargest(top_n, 'Not Overdue').copy()
                plot_data['Overdue'] = 0
            title = f"{title_suffix} Vendors (Not Overdue Only)"

        # === Override with Selected Vendor ===
        if selected_vendor != "Top N":
            plot_data = summary[summary['Vendor_Name'] == selected_vendor]
            title = f"{selected_vendor}"

        # === Melt for Plotting ===
        plot_df = plot_data.melt(id_vars='Vendor_Name', value_vars=['Overdue', 'Not Overdue'],
                                 var_name='Type', value_name='Amount')
        plot_df = plot_df[plot_df['Amount'] > 0]

        # === Bar Chart ===
        fig = px.bar(plot_df, x='Amount', y='Vendor_Name', color='Type',
                     orientation='h', title=title,
                     color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
                     height=max(600, len(plot_df) * 40))

        # Add total labels
        totals = plot_df.groupby('Vendor_Name')['Amount'].sum()
        for vendor, total in totals.items():
            fig.add_annotation(x=total, y=vendor, text=f'€{total:,.0f}',
                               showarrow=False, xanchor='left', font=dict(size=14, color='white'))

        fig.update_layout(
            xaxis_title="Amount (€)", yaxis_title="Vendor", barmode='stack',
            legend_title="Status", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=180, r=60, t=80, b=60)
        )

        chart = st.plotly_chart(fig, use_container_width=True, key="chart", on_select="rerun")

        # === Click to Show Raw Invoices ===
        if chart.selection and chart.selection['points']:
            point = chart.selection['points'][0]
            st.session_state.clicked_vendor = point['y']

        if st.session_state.clicked_vendor:
            vendor = st.session_state.clicked_vendor
            st.subheader(f"Raw Invoices: {vendor}")

            # Apply same filters
            mask = (df['Vendor_Name'] == vendor)
            if status_filter == "Overdue Only":
                mask &= df['Overdue']
            elif status_filter == "Not Overdue Only":
                mask &= ~df['Overdue']

            details = df[mask].copy()
            details = details[['VAT_ID', 'Due_Date', 'Open_Amount', 'BJ_Alt_Invoice', 'Status', 'Vendor_Email', 'Account_Email']]
            details['Due_Date'] = details['Due_Date'].dt.strftime('%Y-%m-%d')
            details['Open_Amount'] = details['Open_Amount'].map('€{:,.2f}'.format)

            st.dataframe(details, use_container_width=True)

            # Download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                details.to_excel(writer, index=False, sheet_name='Invoices')
            buffer.seek(0)
            st.download_button("Download This Vendor", data=buffer,
                               file_name=f"{vendor.replace(' ', '_')}_invoices.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        else:
            st.info("**Click any bar** to view raw invoice details.")

        # === Export All ===
        st.markdown("---")
        st.subheader("Export Raw Data")

        def export_df(data, name):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                data.to_excel(writer, sheet_name='Raw_Data', index=False, startrow=1, header=False)
                worksheet = writer.sheets['Raw_Data']
                workbook = writer.book
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1f4e79', 'font_color': 'white', 'border': 1})
                for i, col in enumerate(data.columns):
                    worksheet.write(0, i, col, header_fmt)
                worksheet.set_column('C:C', 15, workbook.add_format({'num_format': '€#,##0.00'}))
                worksheet.set_column('B:B', 12, workbook.add_format({'num_format': 'dd/mm/yyyy'}))
                worksheet.freeze_panes(1, 0)
            buffer.seek(0)
            return buffer

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("All Open", data=export_df(df, "all"), file_name="All_Open_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c2:
            st.download_button("All Overdue", data=export_df(df[df['Overdue']], "overdue"), file_name="All_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c3:
            st.download_button("All Not Overdue", data=export_df(df[~df['Overdue']], "not"), file_name="All_Not_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.stop()
else:
    st.info("Upload your Excel file to begin.")
