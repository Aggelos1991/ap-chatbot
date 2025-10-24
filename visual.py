# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- Page Config ---
st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Upload Excel → Top 10 Vendors → Click or Select → See Details + Emails**")

# --- Upload ---
uploaded_file = st.file_uploader("Upload Excel (A=Vendor, B=VAT, C=Due Date, G=Amount, AD=Vendor Email, AE=Account Email)", type=['xlsx'])

if uploaded_file:
    try:
        # Read Excel
        df = pd.read_excel(uploaded_file, header=None)

        if df.shape[1] < 31:
            st.error("Excel must have at least 31 columns (A to AE).")
            st.stop()

        # Map columns
        df.columns = [f"C{i}" for i in range(df.shape[1])]
        df = df.rename(columns={
            'C0':  'Vendor_Name',     # A
            'C1':  'VAT_ID',          # B
            'C2':  'Due_Date',        # C
            'C6':  'Open_Amount',     # G
            'C29': 'Vendor_Email',    # AD
            'C30': 'Account_Email'    # AE
        })

        df = df[['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']].copy()

        # Convert
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open invoices found.")
            st.stop()

        # Overdue
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # Summary
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
            vendors = ["Top 10"] + sorted(df['Vendor_Name'].unique().tolist())
            selected = st.selectbox("Vendor", vendors)

        # Apply filter
        plot_df = summary.copy()
        if status_filter == "Overdue Only":
            plot_df['Not_Overdue_Amount'] = 0
        elif status_filter == "Not Overdue Only":
            plot_df['Overdue_Amount'] = 0

        if selected != "Top 10":
            plot_df = plot_df[plot_df['Vendor_Name'] == selected]
            title = f"{selected} - Open Items"
        else:
            plot_df = top10
            title = "Top 10 Vendors"

        # Chart data
        plot_df = plot_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue_Amount', 'Not_Overdue_Amount'],
            var_name='Type',
            value_name='Amount'
        ).replace({'Overdue_Amount': 'Overdue', 'Not_Overdue_Amount': 'Not Overdue'})

        # Bar Chart
        fig = px.bar(
            plot_df,
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
        fig.update_layout(xaxis_title="Amount", yaxis_title="Vendor", legend_title="Status")

        # Show chart
        chart = st.plotly_chart(fig, use_container_width=True, key="main_chart")
        click = st.session_state.get("main_chart", {}).get("clickData")

        # Show details
        show_vendor = selected if selected != "Top 10" else None
        if click:
            show_vendor = click['points'][0]['y']

        if show_vendor:
            st.subheader(f"Details: {show_vendor}")
            det = df[df['Vendor_Name'] == show_vendor][[
                'VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Vendor_Email', 'Account_Email'
            ]].copy()
            det['Due_Date'] = det['Due_Date'].dt.strftime('%Y-%m-%d')
            det['Open_Amount'] = det['Open_Amount'].map('${:,.2f}'.format)
            st.dataframe(det, use_container_width=True)

            # Export
            buf = io.BytesIO()
            det.to_excel(buf, index=False, engine='xlsxwriter')
            buf.seek(0)
            st.download_button(
                "Download Excel",
                data=buf,
                file_name=f"{show_vendor.replace(' ', '_')}_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Click a bar or select a vendor to see details.")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload your Excel file to start.")
