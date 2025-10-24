# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# --- Page Setup ---
st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Upload Excel → View Top 10 Vendors → Click or Select → See Details + Emails**")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload Excel File (A=Vendor, B=VAT, C=Due Date, G=Amount, AD=Vendor Email, AE=Account Email)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        # Read Excel without assuming header
        df = pd.read_excel(uploaded_file, header=None)

        # Check for minimum columns (AE is column 31)
        if df.shape[1] < 31:
            st.error("Error: Excel must have at least 31 columns (A to AE).")
            st.stop()

        # Map columns by position
        df.columns = [f"Col_{i}" for i in range(df.shape[1])]
        df = df.rename(columns={
            df.columns[0]:  'Vendor_Name',     # Column A
            df.columns[1]:  'VAT_ID',          # Column B
            df.columns[2]:  'Due_Date',        # Column C
            df.columns[6]:  'Open_Amount',     # Column G
            df.columns[29]: 'Vendor_Email',    # Column AD
            df.columns[30]: 'Account_Email'    # Column AE
        })

        # Keep only needed columns
        df = df[['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']].copy()

        # Convert types
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')

        # Drop invalid rows
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])

        # Overdue logic
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # Filter open items only
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open invoices found in the file.")
            st.stop()

        # --- Group by Vendor ---
        summary = df.groupby('Vendor_Name').agg(
            Total=('Open_Amount', 'sum'),
            Overdue_Amount=('Open_Amount', lambda x: x[df.loc[x.index, 'Overdue']]),
            Not_Overdue_Amount=('Open_Amount', lambda x: x[~df.loc[x.index, 'Overdue']])
        ).reset_index()
        summary['Not_Overdue_Amount'] = summary['Total'] - summary['Overdue_Amount']
        top10 = summary.nlargest(10, 'Total')

        # --- Filters ---
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"])
        with col2:
            vendor_list = ["Top 10"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_list)

        # --- Apply Filters ---
        plot_data = summary.copy()
        if status_filter == "Overdue Only":
            plot_data['Not_Overdue_Amount'] = 0
        elif status_filter == "Not Overdue Only":
            plot_data['Overdue_Amount'] = 0

        if selected_vendor != "Top 10":
            plot_data = plot_data[plot_data['Vendor_Name'] == selected_vendor]
            title = f"{selected_vendor} - Open Items"
        else:
            plot_data = top10
            title = "Top 10 Vendors by Open Amount"

        # --- Prepare for Chart ---
        plot_data = plot_data.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue_Amount', 'Not_Overdue_Amount'],
            var_name='Type',
            value_name='Amount'
        )
        plot_data['Type'] = plot_data['Type'].replace({
            'Overdue_Amount': 'Overdue',
            'Not_Overdue_Amount': 'Not Overdue'
        })

        # --- Bar Chart ---
        fig = px.bar(
            plot_data,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title,
            color_discrete_map={'Overdue': '#FF6B6B', 'Not Overdue': '#4ECDC4'},
            text='Amount',
            height=500
        )
        fig.update_traces(texttemplate='$%{text:,.0f}', textposition='inside')
        fig.update_layout(
            xaxis_title="Amount ($)",
            yaxis_title="Vendor",
            legend_title="Status",
            hovermode="y unified"
        )

        # Render chart
        chart = st.plotly_chart(fig, use_container_width=True, key="chart")
        click = st.session_state.get("chart", {}).get("clickData")

        # --- Show Details ---
        show_vendor = selected_vendor if selected_vendor != "Top 10" else None
        if click and 'points' in click:
            show_vendor = click['points'][0]['y']

        if show_vendor:
            st.subheader(f"Details: {show_vendor}")
            details = df[df['Vendor_Name'] == show_vendor].copy()
            details = details[[
                'VAT_ID', 'Due_Date', 'Open_Amount', 'Status',
                'Vendor_Email', 'Account_Email'
            ]]
            details['Due_Date'] = details['Due_Date'].dt.strftime('%Y-%m-%d')
            details['Open_Amount'] = details['Open_Amount'].map('${:,.2f}'.format)

            st.dataframe(details, use_container_width=True)

            # Export
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                details.to_excel(writer, index=False, sheet_name='Invoices')
            buffer.seek(0)
            st.download_button(
                "Download Details (Excel)",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Click a bar or select a vendor to view invoice details.")

    except Exception as e:
        st.error(f"Processing error: {str(e)}")
else:
    st.info("Please upload your Excel file.")
