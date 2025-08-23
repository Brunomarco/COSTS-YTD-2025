import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Cost Analysis Dashboard 2025",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Reduce Plotly default height to speed up rendering
px.defaults.height = 400

# Custom CSS for better styling (simplified for performance)
st.markdown("""
    <style>
    .main {padding: 0rem 1rem;}
    h1 {color: #667eea; font-size: 2.5rem; font-weight: bold;}
    </style>
""", unsafe_allow_html=True)

# Title and description
st.title("📊 Cost Analysis Dashboard 2025")
st.markdown("### Comprehensive analysis of order costs and account performance")
st.markdown("---")

# Exchange rates to EUR (you can update these as needed)
EXCHANGE_RATES = {
    'EUR': 1.0,
    'GBP': 1.17,
    'USD': 0.92,
    'KRW': 0.00069,
    'AUD': 0.60,
    'SGD': 0.68,
    # Add more currencies as needed
}

@st.cache_data
def load_and_process_data(file):
    """Load and process the Excel file"""
    try:
        # Read Excel file
        df = pd.read_excel(file, sheet_name=0)
        
        # Convert date column
        df['ORD DT'] = pd.to_datetime(df['ORD DT'], errors='coerce')
        
        # Function to convert any currency to EUR based on CURR column
        def convert_to_eur(row, column):
            """Convert value to EUR based on the row's CURR column"""
            try:
                value = row[column] if pd.notna(row[column]) else 0
                if value == 0 or pd.isna(value):
                    return 0
                
                # Get the currency for this row
                currency = row['CURR'] if pd.notna(row['CURR']) else 'EUR'
                
                # If currency is not in our rates, assume EUR
                if currency not in EXCHANGE_RATES:
                    st.warning(f"Unknown currency '{currency}' found, treating as EUR")
                    return float(value)
                
                # Convert to EUR using the exchange rate
                return float(value) * EXCHANGE_RATES.get(currency, 1.0)
            except:
                return 0
        
        # Convert all cost columns to EUR based on each row's CURR value
        cost_columns = ['PU COST', 'SHIP COST', 'MAN COST', 'DEL COST', 'Total cost', 'NET', 'TOTAL

# File uploader
uploaded_file = st.file_uploader(
    "Upload your Cost Excel file (Cost YTD 2025.xls)",
    type=['xls', 'xlsx'],
    help="Upload the Excel file with cost data"
)

if uploaded_file is not None:
    # Load data
    df = load_and_process_data(uploaded_file)
    
    if df is not None:
        # Sidebar filters
        st.sidebar.header("🔍 Filters")
        
        # Display exchange rates being used
        with st.sidebar.expander("💱 Exchange Rates to EUR"):
            for currency, rate in EXCHANGE_RATES.items():
                if currency != 'EUR':
                    st.text(f"1 {currency} = {rate:.4f} EUR")
        
        # Account filter
        all_accounts = df['ACCT NM'].dropna().unique()
        selected_accounts = st.sidebar.multiselect(
            "Select Accounts",
            options=all_accounts,
            default=None,
            help="Leave empty to show all accounts"
        )
        
        # Country filter
        all_countries = df['PU CTRY'].dropna().unique()
        selected_countries = st.sidebar.multiselect(
            "Select Countries",
            options=all_countries,
            default=None,
            help="Leave empty to show all countries"
        )
        
        st.sidebar.info("📌 Only showing orders with status: 440-BILLED")
        
        # Apply filters (but don't filter by status since we already filtered for 440-BILLED)
        filtered_df = df.copy()
        if selected_accounts:
            filtered_df = filtered_df[filtered_df['ACCT NM'].isin(selected_accounts)]
        if selected_countries:
            filtered_df = filtered_df[filtered_df['PU CTRY'].isin(selected_countries)]
        
        # Key Metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_orders = len(filtered_df)
            st.metric(
                label="📦 Total Billed Orders",
                value=f"{total_orders:,}",
                delta=f"All orders shown are 440-BILLED"
            )
        
        with col2:
            total_cost = filtered_df['Total cost_EUR'].sum()
            st.metric(
                label="💰 Total Cost (EUR)",
                value=f"€{total_cost:,.2f}",
                delta=f"Avg: €{(total_cost/total_orders if total_orders > 0 else 0):,.2f}"
            )
        
        with col3:
            total_net = filtered_df['NET_EUR'].sum()
            difference = total_net - total_cost
            diff_color = "normal" if difference >= 0 else "inverse"
            st.metric(
                label="📈 Total NET (EUR)",
                value=f"€{total_net:,.2f}",
                delta=f"Diff: €{difference:,.2f}",
                delta_color=diff_color
            )
        
        with col4:
            unique_accounts = filtered_df['ACCT'].nunique()
            st.metric(
                label="👥 Unique Accounts",
                value=f"{unique_accounts:,}",
                delta=f"Active: {filtered_df[filtered_df['Total cost_EUR'] > 0]['ACCT'].nunique()}"
            )
        
        st.markdown("---")
        
        # Create two columns for charts
        col1, col2 = st.columns(2)
        
        with col1:
            # Cost Breakdown Pie Chart
            st.subheader("🍩 Cost Breakdown by Type")
            cost_breakdown = {
                'PU Cost': filtered_df['PU COST_EUR'].sum(),
                'Ship Cost': filtered_df['SHIP COST_EUR'].sum(),
                'Man Cost': filtered_df['MAN COST_EUR'].sum(),
                'Del Cost': filtered_df['DEL COST_EUR'].sum()
            }
            
            fig_pie = px.pie(
                values=list(cost_breakdown.values()),
                names=list(cost_breakdown.keys()),
                color_discrete_sequence=['#667eea', '#764ba2', '#f093fb', '#f5576c'],
                hole=0.4
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            # Order Status Distribution - removed since we're only showing 440-BILLED
            st.subheader("📊 Cost Types Distribution")
            
            # Show distribution of which cost types are present
            cost_presence = {
                'PU Cost': (filtered_df['PU COST_EUR'] > 0).sum(),
                'Ship Cost': (filtered_df['SHIP COST_EUR'] > 0).sum(),
                'Man Cost': (filtered_df['MAN COST_EUR'] > 0).sum(),
                'Del Cost': (filtered_df['DEL COST_EUR'] > 0).sum()
            }
            
            fig_cost_presence = px.bar(
                x=list(cost_presence.values()),
                y=list(cost_presence.keys()),
                orientation='h',
                color=list(cost_presence.values()),
                color_continuous_scale='Viridis',
                labels={'x': 'Number of Orders', 'y': 'Cost Type'}
            )
            fig_cost_presence.update_layout(
                height=400,
                showlegend=False,
                xaxis_title="Number of Orders with this Cost Type",
                yaxis_title=""
            )
            st.plotly_chart(fig_cost_presence, use_container_width=True)
        
        # Top Accounts by Cost
        st.markdown("---")
        st.subheader("🏆 Top 10 Accounts by Total Cost")
        
        account_costs = filtered_df.groupby(['ACCT', 'ACCT NM']).agg({
            'Total cost_EUR': 'sum',
            'NET_EUR': 'sum',
            'ORD#': 'count'
        }).reset_index()
        account_costs.columns = ['Account', 'Account Name', 'Total Cost', 'NET', 'Orders']
        account_costs['Difference'] = account_costs['NET'] - account_costs['Total Cost']
        account_costs = account_costs.sort_values('Total Cost', ascending=False).head(10)
        
        fig_top = px.bar(
            account_costs,
            x='Total Cost',
            y='Account Name',
            orientation='h',
            color='Total Cost',
            color_continuous_scale='Blues',
            text='Total Cost'
        )
        fig_top.update_traces(texttemplate='€%{text:,.0f}', textposition='outside')
        fig_top.update_layout(
            height=500,
            xaxis_title="Total Cost (EUR)",
            yaxis_title="",
            showlegend=False
        )
        st.plotly_chart(fig_top, use_container_width=True)
        
        # Monthly Trend and Country Analysis
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Monthly Cost Trend")
            monthly_data = filtered_df.groupby('Month').agg({
                'Total cost_EUR': 'sum',
                'ORD#': 'count'
            }).reset_index()
            monthly_data.columns = ['Month', 'Total Cost', 'Orders']
            
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(
                x=monthly_data['Month'],
                y=monthly_data['Total Cost'],
                name='Total Cost',
                marker_color='#667eea'
            ))
            fig_trend.add_trace(go.Scatter(
                x=monthly_data['Month'],
                y=monthly_data['Orders'] * (monthly_data['Total Cost'].max() / monthly_data['Orders'].max()),
                name='Orders (scaled)',
                yaxis='y2',
                line=dict(color='#f5576c', width=3)
            ))
            fig_trend.update_layout(
                height=400,
                xaxis_title="Month",
                yaxis_title="Total Cost (EUR)",
                yaxis2=dict(
                    title="Orders",
                    overlaying='y',
                    side='right'
                ),
                hovermode='x unified'
            )
            st.plotly_chart(fig_trend, use_container_width=True)
        
        with col2:
            st.subheader("🌍 Top 10 Countries by Cost")
            country_costs = filtered_df.groupby('PU CTRY')['Total cost_EUR'].sum().sort_values(ascending=False).head(10)
            
            fig_country = px.bar(
                x=country_costs.index,
                y=country_costs.values,
                color=country_costs.values,
                color_continuous_scale='Plasma',
                text=country_costs.values
            )
            fig_country.update_traces(texttemplate='€%{text:,.0f}', textposition='outside')
            fig_country.update_layout(
                height=400,
                xaxis_title="Country",
                yaxis_title="Total Cost (EUR)",
                showlegend=False
            )
            st.plotly_chart(fig_country, use_container_width=True)
        
        # Detailed Account Analysis Table
        st.markdown("---")
        st.subheader("📋 Accounts with Highest Cost Differences")
        
        # Calculate account differences
        account_diff = filtered_df.groupby(['ACCT', 'ACCT NM']).agg({
            'Total cost_EUR': 'sum',
            'NET_EUR': 'sum',
            'ORD#': 'count'
        }).reset_index()
        account_diff.columns = ['Account', 'Account Name', 'Total Cost (EUR)', 'NET (EUR)', 'Orders']
        account_diff['Difference (EUR)'] = account_diff['NET (EUR)'] - account_diff['Total Cost (EUR)']
        account_diff['Diff %'] = (account_diff['Difference (EUR)'] / account_diff['Total Cost (EUR)'] * 100).round(2)
        account_diff = account_diff.sort_values('Difference (EUR)', ascending=False, key=abs)
        
        # Format the columns
        for col in ['Total Cost (EUR)', 'NET (EUR)', 'Difference (EUR)']:
            account_diff[col] = account_diff[col].apply(lambda x: f"€{x:,.2f}")
        account_diff['Diff %'] = account_diff['Diff %'].apply(lambda x: f"{x:.1f}%")
        
        # Display table with conditional formatting
        st.dataframe(
            account_diff.head(15),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Account": st.column_config.TextColumn("Account", width="small"),
                "Account Name": st.column_config.TextColumn("Account Name", width="large"),
                "Orders": st.column_config.NumberColumn("Orders", width="small"),
            }
        )
        
        # Download processed data
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            csv = filtered_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Filtered Data (CSV)",
                data=csv,
                file_name=f"cost_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        with col2:
            # Summary statistics
            st.download_button(
                label="📊 Download Summary Report (CSV)",
                data=account_diff.to_csv(index=False),
                file_name=f"account_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        # Footer
        st.markdown("---")
        st.caption(f"Dashboard generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Data contains {len(df):,} total records")
        
else:
    # Instructions when no file is uploaded
    st.info("👆 Please upload your Excel file to begin the analysis")
    st.markdown("""
    ### Expected Excel Format:
    Your Excel file should contain the following columns:
    - **ORD DT**: Order Date
    - **ACCT**: Account Number
    - **ACCT NM**: Account Name
    - **OFC**: Office
    - **ORD#**: Order Number
    - **PU COST**: Pickup Cost
    - **SHIP COST**: Shipping Cost
    - **MAN COST**: Manufacturing Cost
    - **DEL COST**: Delivery Cost
    - **Total cost**: Total Cost
    - **NET**: Net Amount
    - **CURR**: Currency (EUR/GBP/USD/KRW/AUD/SGD)
    - **INV#**: Invoice Number
    - **TOTAL$**: Total Amount
    - **STATUS**: Order Status
    - **PU CTRY**: Pickup Country
    
    ### Currency Conversion:
    All amounts will be automatically converted to EUR based on the CURR column for each row.
    Current exchange rates to EUR:
    - GBP: 1.17
    - USD: 0.92
    - KRW: 0.00069
    - AUD: 0.60
    - SGD: 0.68
    """)
]
        for col in cost_columns:
            df[f'{col}_EUR'] = df.apply(lambda row: convert_to_eur(row, col), axis=1)
        
        # Fill NaN values with 0 for cost columns
        for col in cost_columns:
            df[f'{col}_EUR'] = df[f'{col}_EUR'].fillna(0)
            df[col] = df[col].fillna(0)
        
        # Extract month and year
        df['Month'] = df['ORD DT'].dt.to_period('M').astype(str)
        
        # Display currency distribution info
        if 'CURR' in df.columns:
            currency_counts = df['CURR'].value_counts()
            st.sidebar.markdown("### 💱 Currency Distribution")
            for curr, count in currency_counts.items():
                if pd.notna(curr):
                    rate = EXCHANGE_RATES.get(curr, 1.0)
                    st.sidebar.text(f"{curr}: {count} orders (rate: {rate})")
        
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

# File uploader
uploaded_file = st.file_uploader(
    "Upload your Cost Excel file (Cost YTD 2025.xls)",
    type=['xls', 'xlsx'],
    help="Upload the Excel file with cost data"
)

if uploaded_file is not None:
    # Load data
    df = load_and_process_data(uploaded_file)
    
    if df is not None:
        # Sidebar filters
        st.sidebar.header("🔍 Filters")
        
        # Account filter
        all_accounts = df['ACCT NM'].dropna().unique()
        selected_accounts = st.sidebar.multiselect(
            "Select Accounts",
            options=all_accounts,
            default=None,
            help="Leave empty to show all accounts"
        )
        
        # Status filter
        all_statuses = df['STATUS'].dropna().unique()
        selected_statuses = st.sidebar.multiselect(
            "Select Status",
            options=all_statuses,
            default=None,
            help="Leave empty to show all statuses"
        )
        
        # Country filter
        all_countries = df['PU CTRY'].dropna().unique()
        selected_countries = st.sidebar.multiselect(
            "Select Countries",
            options=all_countries,
            default=None,
            help="Leave empty to show all countries"
        )
        
        # Apply filters
        filtered_df = df.copy()
        if selected_accounts:
            filtered_df = filtered_df[filtered_df['ACCT NM'].isin(selected_accounts)]
        if selected_statuses:
            filtered_df = filtered_df[filtered_df['STATUS'].isin(selected_statuses)]
        if selected_countries:
            filtered_df = filtered_df[filtered_df['PU CTRY'].isin(selected_countries)]
        
        # Key Metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_orders = len(filtered_df)
            st.metric(
                label="📦 Total Orders",
                value=f"{total_orders:,}",
                delta=f"{len(filtered_df[filtered_df['STATUS'] == '440-BILLED']):,} Billed"
            )
        
        with col2:
            total_cost = filtered_df['Total cost_EUR'].sum()
            st.metric(
                label="💰 Total Cost (EUR)",
                value=f"€{total_cost:,.2f}",
                delta=f"Avg: €{(total_cost/total_orders if total_orders > 0 else 0):,.2f}"
            )
        
        with col3:
            total_net = filtered_df['NET_EUR'].sum()
            st.metric(
                label="📈 Total NET (EUR)",
                value=f"€{total_net:,.2f}",
                delta=f"Diff: €{(total_net - total_cost):,.2f}"
            )
        
        with col4:
            unique_accounts = filtered_df['ACCT'].nunique()
            st.metric(
                label="👥 Unique Accounts",
                value=f"{unique_accounts:,}",
                delta=f"Active: {filtered_df[filtered_df['Total cost_EUR'] > 0]['ACCT'].nunique()}"
            )
        
        st.markdown("---")
        
        # Create two columns for charts
        col1, col2 = st.columns(2)
        
        with col1:
            # Cost Breakdown Pie Chart
            st.subheader("🍩 Cost Breakdown by Type")
            cost_breakdown = {
                'PU Cost': filtered_df['PU COST_EUR'].sum(),
                'Ship Cost': filtered_df['SHIP COST_EUR'].sum(),
                'Man Cost': filtered_df['MAN COST_EUR'].sum(),
                'Del Cost': filtered_df['DEL COST_EUR'].sum()
            }
            
            fig_pie = px.pie(
                values=list(cost_breakdown.values()),
                names=list(cost_breakdown.keys()),
                color_discrete_sequence=['#667eea', '#764ba2', '#f093fb', '#f5576c'],
                hole=0.4
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            # Order Status Distribution
            st.subheader("📊 Order Status Distribution")
            status_counts = filtered_df['STATUS'].value_counts().head(8)
            
            fig_status = px.bar(
                x=status_counts.values,
                y=status_counts.index,
                orientation='h',
                color=status_counts.values,
                color_continuous_scale='Viridis'
            )
            fig_status.update_layout(
                height=400,
                showlegend=False,
                xaxis_title="Number of Orders",
                yaxis_title="Status"
            )
            st.plotly_chart(fig_status, use_container_width=True)
        
        # Top Accounts by Cost
        st.markdown("---")
        st.subheader("🏆 Top 10 Accounts by Total Cost")
        
        account_costs = filtered_df.groupby(['ACCT', 'ACCT NM']).agg({
            'Total cost_EUR': 'sum',
            'NET_EUR': 'sum',
            'ORD#': 'count'
        }).reset_index()
        account_costs.columns = ['Account', 'Account Name', 'Total Cost', 'NET', 'Orders']
        account_costs['Difference'] = account_costs['NET'] - account_costs['Total Cost']
        account_costs = account_costs.sort_values('Total Cost', ascending=False).head(10)
        
        fig_top = px.bar(
            account_costs,
            x='Total Cost',
            y='Account Name',
            orientation='h',
            color='Total Cost',
            color_continuous_scale='Blues',
            text='Total Cost'
        )
        fig_top.update_traces(texttemplate='€%{text:,.0f}', textposition='outside')
        fig_top.update_layout(
            height=500,
            xaxis_title="Total Cost (EUR)",
            yaxis_title="",
            showlegend=False
        )
        st.plotly_chart(fig_top, use_container_width=True)
        
        # Monthly Trend and Country Analysis
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Monthly Cost Trend")
            monthly_data = filtered_df.groupby('Month').agg({
                'Total cost_EUR': 'sum',
                'ORD#': 'count'
            }).reset_index()
            monthly_data.columns = ['Month', 'Total Cost', 'Orders']
            
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(
                x=monthly_data['Month'],
                y=monthly_data['Total Cost'],
                name='Total Cost',
                marker_color='#667eea'
            ))
            fig_trend.add_trace(go.Scatter(
                x=monthly_data['Month'],
                y=monthly_data['Orders'] * (monthly_data['Total Cost'].max() / monthly_data['Orders'].max()),
                name='Orders (scaled)',
                yaxis='y2',
                line=dict(color='#f5576c', width=3)
            ))
            fig_trend.update_layout(
                height=400,
                xaxis_title="Month",
                yaxis_title="Total Cost (EUR)",
                yaxis2=dict(
                    title="Orders",
                    overlaying='y',
                    side='right'
                ),
                hovermode='x unified'
            )
            st.plotly_chart(fig_trend, use_container_width=True)
        
        with col2:
            st.subheader("🌍 Top 10 Countries by Cost")
            country_costs = filtered_df.groupby('PU CTRY')['Total cost_EUR'].sum().sort_values(ascending=False).head(10)
            
            fig_country = px.bar(
                x=country_costs.index,
                y=country_costs.values,
                color=country_costs.values,
                color_continuous_scale='Plasma',
                text=country_costs.values
            )
            fig_country.update_traces(texttemplate='€%{text:,.0f}', textposition='outside')
            fig_country.update_layout(
                height=400,
                xaxis_title="Country",
                yaxis_title="Total Cost (EUR)",
                showlegend=False
            )
            st.plotly_chart(fig_country, use_container_width=True)
        
        # Detailed Account Analysis Table
        st.markdown("---")
        st.subheader("📋 Accounts with Highest Cost Differences")
        
        # Calculate account differences
        account_diff = filtered_df.groupby(['ACCT', 'ACCT NM']).agg({
            'Total cost_EUR': 'sum',
            'NET_EUR': 'sum',
            'ORD#': 'count'
        }).reset_index()
        account_diff.columns = ['Account', 'Account Name', 'Total Cost (EUR)', 'NET (EUR)', 'Orders']
        account_diff['Difference (EUR)'] = account_diff['NET (EUR)'] - account_diff['Total Cost (EUR)']
        account_diff['Diff %'] = (account_diff['Difference (EUR)'] / account_diff['Total Cost (EUR)'] * 100).round(2)
        account_diff = account_diff.sort_values('Difference (EUR)', ascending=False, key=abs)
        
        # Format the columns
        for col in ['Total Cost (EUR)', 'NET (EUR)', 'Difference (EUR)']:
            account_diff[col] = account_diff[col].apply(lambda x: f"€{x:,.2f}")
        account_diff['Diff %'] = account_diff['Diff %'].apply(lambda x: f"{x:.1f}%")
        
        # Display table with conditional formatting
        st.dataframe(
            account_diff.head(15),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Account": st.column_config.TextColumn("Account", width="small"),
                "Account Name": st.column_config.TextColumn("Account Name", width="large"),
                "Orders": st.column_config.NumberColumn("Orders", width="small"),
            }
        )
        
        # Download processed data
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            csv = filtered_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Filtered Data (CSV)",
                data=csv,
                file_name=f"cost_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        with col2:
            # Summary statistics
            st.download_button(
                label="📊 Download Summary Report (CSV)",
                data=account_diff.to_csv(index=False),
                file_name=f"account_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        # Footer
        st.markdown("---")
        st.caption(f"Dashboard generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Data contains {len(df):,} total records")
        
else:
    # Instructions when no file is uploaded
    st.info("👆 Please upload your Excel file to begin the analysis")
    st.markdown("""
    ### Expected Excel Format:
    Your Excel file should contain the following columns:
    - **ORD DT**: Order Date
    - **ACCT**: Account Number
    - **ACCT NM**: Account Name
    - **OFC**: Office
    - **ORD#**: Order Number
    - **PU COST**: Pickup Cost
    - **SHIP COST**: Shipping Cost
    - **MAN COST**: Manufacturing Cost
    - **DEL COST**: Delivery Cost
    - **Total cost**: Total Cost
    - **NET**: Net Amount
    - **CURR**: Currency (EUR/GBP)
    - **INV#**: Invoice Number
    - **TOTAL$**: Total Amount
    - **STATUS**: Order Status
    - **PU CTRY**: Pickup Country
    """)
