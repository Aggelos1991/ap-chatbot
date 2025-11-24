import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import hashlib
from datetime import date, timedelta

# --- CONFIGURATION & CONSTANTS ---
TODAY = pd.Timestamp.now(tz='Europe/Athens').date()

# Define the expected sheet names and column indices (based on the original code's logic)
CONFIG = {
    'MAIN_SHEET': 'Outstanding Invoices IB',
    'REF_SHEET': 'VR CHECK_Special vendors list',
    'COUNTRY_SHEET': 'Vendors',
    'MAIN_COLS_INDICES': [0, 1, 4, 6, 29, 30, 31, 33, 35, 39],
    'MAIN_COL_NAMES': [
        'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
        'Vendor_Email', 'Account_Email', 'Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN'
    ],
    'REF_COLS': ['Vendor_TaxID', 'Col_B', 'Col_C', 'Col_D', 'Col_E', 'Vendor_Category'],
    'COUNTRY_COLS': ['Vendor_TaxID', 'Col_B', 'Col_C', 'Col_D', 'Col_E', 'Col_F', 'Country']
}


# --- DATA LOADING AND CLEANING ---

@st.cache_data(show_spinner="Processing the ledger... binding the invoices...")
def load_data(uploaded_file):
    """Loads, cleans, and merges data from the uploaded Excel file."""
    
    with pd.ExcelFile(uploaded_file) as xls:
        
        # 1. Load Main Data
        if CONFIG['MAIN_SHEET'] not in xls.sheet_names:
            st.error(f"Sheet '{CONFIG['MAIN_SHEET']}' not found. Please check your file.")
            return None
        df_raw = pd.read_excel(xls, sheet_name=CONFIG['MAIN_SHEET'], header=None)

        # Header Detection
        header_row_index = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row_index.empty:
            st.error("Header 'VENDOR' not found in column A. Cannot determine data start row.")
            return None
        start_row = header_row_index[0] + 1
        
        # Extract main data and columns
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)
        df = df.iloc[:, CONFIG['MAIN_COLS_INDICES']]
        df.columns = CONFIG['MAIN_COL_NAMES']

        # Find BS (Block Status) / BA (Business Area) Columns
        headers = df_raw.iloc[header_row_index[0]].astype(str).str.strip().tolist()
        col_map = {h.upper().strip(): i for i, h in enumerate(headers)}
        # Heuristic search for BS and BA columns based on header names
        bs_idx = next((i for name, i in col_map.items() if "BS" in name and "FUNC" not in name), 50)
        ba_idx = next((i for name, i in col_map.items() if "BA" in name), 51)
        df['Col_BS'] = df_raw.iloc[start_row:, bs_idx].astype(str).str.strip()
        df['Col_BA'] = df_raw.iloc[start_row:, ba_idx].astype(str).str.strip()

        # 2. Load Reference Sheets
        
        # Vendor Category (Ref Sheet)
        if CONFIG['REF_SHEET'] in xls.sheet_names:
            df_ref = pd.read_excel(xls, sheet_name=CONFIG['REF_SHEET'], usecols=list(range(len(CONFIG['REF_COLS']))), header=None)
            df_ref.columns = CONFIG['REF_COLS']
            df_ref['Vendor_TaxID'] = df_ref['Vendor_TaxID'].astype(str).str.strip().str.upper()
        else:
            df_ref = pd.DataFrame(columns=['Vendor_TaxID', 'Vendor_Category'])
            st.warning(f"Sheet '{CONFIG['REF_SHEET']}' not found. Vendor categories will be 'Uncategorized'.")

        # Country Lookup
        if CONFIG['COUNTRY_SHEET'] in xls.sheet_names:
            df_country = pd.read_excel(xls, sheet_name=CONFIG['COUNTRY_SHEET'], usecols=list(range(len(CONFIG['COUNTRY_COLS']))), header=None)
            df_country.columns = CONFIG['COUNTRY_COLS']
            df_country['Vendor_TaxID'] = df_country['Vendor_TaxID'].astype(str).str.strip().str.upper()
        else:
            df_country = pd.DataFrame(columns=['Vendor_TaxID', 'Country'])
            st.warning(f"Sheet '{CONFIG['COUNTRY_SHEET']}' not found. Country classification may be incomplete.")


    # --- DATA CLEANING ---
    df = df.dropna(how="all").dropna(subset=['Vendor_Name'])
    
    # Remove aggregated rows
    bad_patterns = r"(?i)total|saldo|asiento|header|proveedor|unnamed|vendor|facturas|periodo|sum|importe|grand total"
    df = df[~df['Vendor_Name'].astype(str).str.contains(bad_patterns, na=False)]
    
    # Data type conversion and cleaning
    df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce').dt.date
    df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
    df = df.dropna(subset=['Due_Date', 'Open_Amount'])
    df = df[df['Open_Amount'] > 0]
    df['VAT_ID_clean'] = df['VAT_ID'].astype(str).str.strip().str.upper()

    # --- MERGE DATA ---
    
    # Merge Vendor Type
    if not df_ref.empty:
        df = df.merge(df_ref[['Vendor_TaxID', 'Vendor_Category']],
                      left_on='VAT_ID_clean', right_on='Vendor_TaxID', how='left')
        df['Vendor_Type'] = df['Vendor_Category'].fillna("Uncategorized")
    else:
        df['Vendor_Type'] = "Uncategorized"

    # Merge Country Info
    df = df.merge(df_country[['Vendor_TaxID', 'Country']],
                  left_on='VAT_ID_clean', right_on='Vendor_TaxID', how='left', suffixes=('_x', '_y'))
    
    # Define Country Type
    df['Country_Type'] = df['Country'].fillna("Unknown").astype(str).apply(
        lambda x: "Spain" if "spain" in x.lower()
        else "Foreign" if x.strip() != "" and x.lower() != "unknown"
        else "Unknown"
    )

    # --- NORMALIZE STATUS COLUMNS ---
    
    # Normalize Blocked for Payment (BFP)
    def normalize_bs(x):
        x = str(x).strip().upper()
        if x in ["", "OK", "FREE", "0", "FREE FOR PAYMENT", "0.0"]:
            return "Free for Payment"
        elif "BLOCK" in x or x in ["1", "BFP", "1.0"]:
            return "Blocked for Payment"
        return "Other Block Status"
    df['Col_BS'] = df['Col_BS'].apply(normalize_bs)

    # Normalize 'Yes' Filters (AF, AH, AJ, AN)
    for col in ['Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN']:
        df[col] = df[col].fillna("").astype(str).str.strip().str.lower()
        df[col] = df[col].apply(lambda x: 'Yes' if x in ['yes', 'y'] else 'No')


    # --- CORE METRICS ---
    df['Overdue'] = df['Due_Date'] < TODAY
    df['Status'] = np.where(df['Overdue'], 'Overdue', 'Not Overdue')
    # 🔧 Force datetime conversion for required columns
    for col in ['Due_Date', 'Overdue', 'Invoice Date', 'Posting Date']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    df['Days_Overdue'] = np.where(
        df['Overdue'].notna(),
        (TODAY - df['Due_Date']).dt.days,
        0
    )

    # Final data selection for dashboard
    df = df[['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Days_Overdue', 
             'Vendor_Email', 'Account_Email', 'Col_AF', 'Col_AH', 'Col_AJ', 'Col_AN',
             'Vendor_Type', 'Col_BS', 'Country_Type', 'Col_BA']].rename(columns={'Col_BA': 'Business_Area'})
    
    return df


# --- DASHBOARD FUNCTIONS ---

def display_kpis(df):
    """Displays key metrics in three columns."""
    
    overdue_df = df[df['Status'] == 'Overdue']
    
    total_overdue = overdue_df['Open_Amount'].sum()
    total_vendors = overdue_df['Vendor_Name'].nunique()
    avg_days_overdue = overdue_df['Days_Overdue'].replace(0, np.nan).mean()
    
    col1, col2, col3 = st.columns(3)
    
    # KPI 1: Total Overdue Amount
    col1.metric(
        label="💰 Total Overdue Amount", 
        value=f"€{total_overdue:,.0f}", 
        delta=f"from {df['Open_Amount'].sum():,.0f} Total Open"
    )
    
    # KPI 2: Unique Overdue Vendors
    col2.metric(
        label="👥 Unique Overdue Vendors", 
        value=f"{total_vendors:,}", 
        delta=f"{df['Vendor_Name'].nunique()} Total Vendors"
    )

    # KPI 3: Average Days Past Due (for overdue invoices)
    col3.metric(
        label="⏱️ Average Days Past Due", 
        value=f"{avg_days_overdue:.1f} days" if not pd.isna(avg_days_overdue) else "N/A"
    )
    st.markdown("---")


def plot_vendor_summary(df, status_filter, vendor_select, country_choice):
    """Generates the stacked bar chart for vendor summary."""
    
    # Aggregate data
    summary = (
        df.groupby(['Vendor_Name', 'Status'], as_index=False)['Open_Amount']
          .sum()
          .pivot(index='Vendor_Name', columns='Status', values='Open_Amount')
          .fillna(0)
          .reset_index()
    )
    
    # Ensure both status columns exist
    for col in ['Overdue', 'Not Overdue']:
        if col not in summary.columns: summary[col] = 0
    
    summary['Total'] = summary['Overdue'] + summary['Not Overdue']

    # Top N / Specific Vendor Logic
    if "Top" in vendor_select:
        top_n = int(vendor_select.split(" ")[1])
        if status_filter == "All Open":
            base_df = summary.nlargest(top_n, 'Total')
        elif status_filter == "Overdue Only":
            base_df = summary.nlargest(top_n, 'Overdue')
        else: # Not Overdue Only
            base_df = summary.nlargest(top_n, 'Not Overdue')
    else:
        base_df = summary[summary['Vendor_Name'] == vendor_select]

    # Melt for plotting
    plot_df = base_df.melt(id_vars='Vendor_Name',
                           value_vars=['Overdue', 'Not Overdue'],
                           var_name='Type', value_name='Amount').query("Amount>0")

    # Plotly Figure
    fig = px.bar(
        plot_df.sort_values(by='Amount', ascending=True), # Sort for better display
        x='Amount', y='Vendor_Name', color='Type',
        orientation='h', 
        color_discrete_map={'Overdue': '#E53935', 'Not Overdue': '#1E88E5'}, # Richer colors
        title=f"Vendor Balances: {vendor_select} ({status_filter}) — {country_choice}",
        height=max(500, len(plot_df['Vendor_Name'].unique()) * 35 + 150)
    )

    # Custom Layout
    fig.update_layout(
        xaxis_title="Amount (€)", yaxis_title="Vendor", legend_title="Status",
        barmode='stack', 
        plot_bgcolor='rgba(255,255,255,0.05)', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color="#FFFFFF"),
        margin=dict(l=150, r=50, t=80, b=50)
    )
    
    # Force key update for selection reset
    data_hash = hashlib.md5(pd.util.hash_pandas_object(base_df, index=True).values).hexdigest()[:8]
    chart_key = f"bar_chart_{vendor_select}_{status_filter}_{country_choice}_{data_hash}"
    
    # Display and handle click selection
    chart = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key=chart_key)
    
    # Extract selected vendor name from chart selection if available
    selected_vendor = None
    if chart.selection and chart.selection.get('points'):
        selected_vendor = chart.selection['points'][0].get('y')
    
    return selected_vendor


def main_app():
    """Main Streamlit application logic."""
    
    st.set_page_config(page_title="Overdue Invoices", layout="wide")
    st.markdown("""
    <style>
        .stButton>button {
            background-color: #333333;
            color: #d4af37;
            border: 1px solid #d4af37;
        }
        .stButton>button:hover {
            background-color: #d4af37;
            color: #000000;
        }
        h1 {
            text-align:center;
            font-family: 'Cinzel Decorative', serif;
            font-size: 38px;
            color: goldenrod;
            text-shadow: 2px 2px 6px #000000;
        }
        span.subtitle {
            font-size: 24px; 
            color: #d4af37;
            display: block;
            margin-top: -10px;
        }
        .stContainer, .stPlotlyChart {
            background-color: #222222;
            padding: 15px;
            border-radius: 10px;
        }
    </style>
    <link href="https://fonts.googleapis.com/css2?family=Cinzel+Decorative:wght@700&display=swap" rel="stylesheet">
    <h1>
    💍 One Invoice to rule them all,<br>
    One Invoice to seek and find them,<br>
    One Invoice to bring them all,<br>
    And in the Ledger bind them ⚔️<br>
    <span class='subtitle'>In the realm of Overdues, where the balances lie. 📜</span>
    </h1>
    """, unsafe_allow_html=True)
    
    # --- FILE UPLOADER ---
    uploaded_file = st.file_uploader("Upload Excel file (Must contain sheets: 'Outstanding Invoices IB', 'Vendors', and 'VR CHECK_Special vendors list')", type=['xlsx'])
    
    if not uploaded_file:
        st.info("Upload your Excel file to begin the quest for overdue invoices.")
        return

    # --- LOAD DATA ---
    df = load_data(uploaded_file)
    if df is None or df.empty:
        st.warning("No valid data found after cleaning and filtering. Please check your Excel file structure.")
        return

    st.success(f"Data loaded successfully! {len(df):,} invoices processed.")
    
    # --- FILTERS SECTION ---
    st.header("🔍 The Filters of Middle-earth")
    
    # Define a container for all filters
    with st.container(border=True):
        
        # 1. Row: Country and Date Range
        c_country, c_due_date = st.columns([1, 2])
        
        with c_country:
            country_choice = st.radio(
                "Select Country Group", 
                ["All", "Spain", "Foreign"], 
                horizontal=True, 
                index=0
            )

        with c_due_date:
            min_date = df['Due_Date'].min()
            max_date = df['Due_Date'].max()
            date_range = st.slider(
                "Filter Due Date Range",
                value=(min_date if pd.notna(min_date) else TODAY, max_date if pd.notna(max_date) else TODAY),
                format="YYYY/MM/DD"
            )

        # Apply Country and Date Range Filters
        if country_choice != "All":
            df = df[df['Country_Type'] == country_choice]
        df = df[(df['Due_Date'] >= date_range[0]) & (df['Due_Date'] <= date_range[1])]

        # 2. Row: Advanced Status Filters (Expander)
        with st.expander("🛡️ Advanced Payment Status Filters"):
            
            c_yes_filter, c_aj_filter = st.columns(2)
            with c_yes_filter:
                apply_yes = st.checkbox("Filter Col_AF / Col_AH / Col_AN to 'Yes' only", value=True)
            with c_aj_filter:
                aj_yes_only = st.checkbox("Filter Col_AJ = 'Yes' only", value=False)

            if apply_yes:
                for col in ['Col_AF', 'Col_AH', 'Col_AN']:
                    df = df[df[col] == 'Yes']
            
            if aj_yes_only:
                df = df[df['Col_AJ'] == 'Yes']

            c_bt, c_bs = st.columns(2)
            
            # Vendor Type Filter
            with c_bt:
                bt_values = sorted(df['Vendor_Type'].unique())
                selected_bt = st.multiselect("Exceptions / Priority Vendors (Vendor Type)", bt_values, default=bt_values)
                df = df[df['Vendor_Type'].isin(selected_bt)]

            # BFP Status Filter
            with c_bs:
                bs_values = sorted(df['Col_BS'].unique())
                selected_bs = st.multiselect("BFP Status (BS)", bs_values, default=bs_values)
                df = df[df['Col_BS'].isin(selected_bs)]

    if df.empty:
        st.warning("No invoices match the current filter selection.")
        return

    # --- KPIS ---
    display_kpis(df)

    # --- CHART FILTERS & PLOT ---
    st.header("📊 The Balance Scroll")
    
    # Define vendor and status filters for the chart
    c_status, c_vendor_select = st.columns(2)
    with c_status:
        status_filter = st.selectbox("Chart Status View", ["All Open", "Overdue Only", "Not Overdue Only"])
    with c_vendor_select:
        vendor_options = ["Top 200", "Top 100", "Top 30", "Top 20"] + sorted(df['Vendor_Name'].unique())
        vendor_select = st.selectbox("Select Vendor Group or Specific Vendor", vendor_options)

    # Plot chart and get the clicked vendor (if any)
    clicked_vendor = plot_vendor_summary(df, status_filter, vendor_select, country_choice)

    # --- RAW DATA TABLE ---
    st.header("📜 Raw Invoice Ledger")

    filtered_df = df.copy()
    
    # Apply chart status filter to the table
    if status_filter == "Overdue Only":
        filtered_df = filtered_df[filtered_df['Status'] == "Overdue"]
    elif status_filter == "Not Overdue Only":
        filtered_df = filtered_df[filtered_df['Status'] == "Not Overdue"]

    # Apply chart click filter to the table
    if clicked_vendor:
        filtered_df = filtered_df[filtered_df['Vendor_Name'] == clicked_vendor]
        st.subheader(f"Invoices for Selected Vendor: **{clicked_vendor}**")
    elif 'Top' not in vendor_select:
        filtered_df = filtered_df[filtered_df['Vendor_Name'] == vendor_select]
        st.subheader(f"Invoices for Selected Vendor: **{vendor_select}**")
    else:
        # If showing Top N, aggregate to show the top N vendors in the raw data as well.
        top_n_vendors = plot_vendor_summary(df, status_filter, vendor_select, country_choice).reset_index()['Vendor_Name'].tolist()
        filtered_df = filtered_df[filtered_df['Vendor_Name'].isin(top_n_vendors)]
        st.subheader(f"Raw Data for {vendor_select} matching filters")
    
    # Prepare table for display
    show = filtered_df[['Vendor_Name','VAT_ID','Due_Date','Open_Amount','Status','Days_Overdue',
                        'Vendor_Type','Country_Type', 'Business_Area', 'Col_BS',
                        'Col_AF','Col_AH','Col_AJ','Col_AN',
                        'Vendor_Email','Account_Email']].copy()

    show['Due_Date'] = pd.to_datetime(show['Due_Date']).dt.strftime("%Y-%m-%d")
    show['Open_Amount'] = show['Open_Amount'].map('€{:,.2f}'.format)
    
    if not show.empty:
        st.dataframe(show, use_container_width=True, height=400,
                     column_order=('Vendor_Name','Due_Date','Open_Amount','Status','Days_Overdue',
                                   'Vendor_Type', 'Country_Type', 'Business_Area', 'Col_BS',
                                   'VAT_ID', 'Col_AF','Col_AH','Col_AJ','Col_AN',
                                   'Vendor_Email','Account_Email'))
    else:
        st.info("No raw invoices found for the current selection.")

    # --- EMAILS ---
    st.header("📧 The Messenger's Scroll")
    
    emails = pd.concat([filtered_df['Vendor_Email'], filtered_df['Account_Email']], ignore_index=True)
    emails = emails.dropna().astype(str)
    emails = emails[emails.str.contains('@')].str.lower().unique()
    
    lang = country_choice if country_choice in ["Spain", "Foreign"] else "Mixed"
    
    st.code(f"Unique Emails ({len(emails)} for {lang} vendors):", language='text')
    st.text_area(
        "Copy and Paste into your email client (Ctrl + C to copy)", 
        ", ".join(sorted(emails)), 
        height=150
    )
    st.success(f"{len(emails)} unique emails ready for communication.")

if __name__ == "__main__":
    try:
        main_app()
    except Exception as e:
        # Catch any unhandled errors gracefully
        st.error(f"An unexpected error occurred during execution: {e}")
        st.exception(e)
