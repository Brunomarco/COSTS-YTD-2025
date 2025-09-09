import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from plotly.subplots import make_subplots

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

st.markdown("""
    <style>
    [data-testid="stMetricValue"] { white-space: nowrap; overflow: visible; }
    [data-testid="stMetricDelta"] { white-space: nowrap; overflow: visible; }
    </style>
""", unsafe_allow_html=True)

# Title and description
st.title("Cost Analysis Dashboard 2025")
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
        # Read Excel file with proper handling of thousand separators
        df = pd.read_excel(file, sheet_name=0, thousands=',')
        
        # Clean column names (remove extra spaces)
        df.columns = df.columns.str.strip()
        
        # Filter only 440-BILLED status rows
        df = df[df['STATUS'] == '440-BILLED'].copy()
        
        # Convert date column
        df['ORD DT'] = pd.to_datetime(df['ORD DT'], errors='coerce')
        
        # Function to clean and convert numeric values
        def clean_numeric(value):
            """Clean numeric values that might have commas or be strings"""
            if pd.isna(value) or value == '' or value == ' ':
                return 0
            if isinstance(value, str):
                # Remove commas and spaces, then convert to float
                value = value.replace(',', '').replace(' ', '')
                try:
                    return float(value) if value else 0
                except:
                    return 0
            return float(value)
        
        # Clean all cost columns first
        cost_columns = ['PU COST', 'SHIP COST', 'MAN COST', 'DEL COST', 'Total cost', 'NET', 'TOTAL$']
        for col in cost_columns:
            if col in df.columns:
                df[col] = df[col].apply(clean_numeric)
        
        # Function to convert any currency to EUR based on CURR column
        def convert_to_eur(row, column):
            """Convert value to EUR based on the row's CURR column"""
            try:
                value = row[column] if column in row and pd.notna(row[column]) else 0
                if value == 0:
                    return 0
                
                # Get the currency for this row
                currency = str(row['CURR']).strip().upper() if pd.notna(row['CURR']) else 'EUR'
                
                # If currency is not in our rates, assume EUR
                if currency not in EXCHANGE_RATES:
                    if currency and currency != 'NAN':
                        st.sidebar.warning(f"Unknown currency '{currency}' found, treating as EUR")
                    return float(value)
                
                # Convert to EUR using the exchange rate
                return float(value) * EXCHANGE_RATES.get(currency, 1.0)
            except Exception as e:
                return 0
        
        # Convert all cost columns to EUR based on each row's CURR value
        for col in cost_columns:
            if col in df.columns:
                df[f'{col}_EUR'] = df.apply(lambda row: convert_to_eur(row, col), axis=1)
            else:
                df[f'{col}_EUR'] = 0
        
        # Fill NaN values with 0 for EUR columns
        for col in cost_columns:
            if f'{col}_EUR' in df.columns:
                df[f'{col}_EUR'] = df[f'{col}_EUR'].fillna(0)
        
        # Extract month and year
        df['Month'] = df['ORD DT'].dt.to_period('M').astype(str) if 'ORD DT' in df.columns else 'Unknown'
        
        # Display data info in sidebar
        st.sidebar.markdown("### Data Overview")
        st.sidebar.text(f"Total Billed Orders: {len(df)}")
        
        # Display currency distribution info
        if 'CURR' in df.columns:
            currency_counts = df['CURR'].value_counts()
            st.sidebar.markdown("### Currency Distribution")
            for curr, count in currency_counts.items():
                if pd.notna(curr):
                    rate = EXCHANGE_RATES.get(str(curr).strip().upper(), 1.0)
                    st.sidebar.text(f"{curr}: {count} orders (rate: {rate})")
        
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        st.error("Please check that your Excel file has the correct format and columns")
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
        st.sidebar.header("Filters")
        
        # Display exchange rates being used
        with st.sidebar.expander("Exchange Rates to EUR"):
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
        
        st.sidebar.info("Only showing orders with status: 440-BILLED")
        
        # Apply filters
        filtered_df = df.copy()
        if selected_accounts:
            filtered_df = filtered_df[filtered_df['ACCT NM'].isin(selected_accounts)]
        if selected_countries:
            filtered_df = filtered_df[filtered_df['PU CTRY'].isin(selected_countries)]
        
        # Key Metrics
        col1, col2, col3, col4 = st.columns([1.1, 1.4, 1.4, 1.0])
        
        with col1:
            total_orders = len(filtered_df)
            st.metric(
                label="Total Billed Orders",
                value=f"{total_orders:,}",
                delta="All orders shown are 440-BILLED"
            )
        
        with col2:
            total_cost = filtered_df['Total cost_EUR'].sum()
            avg_cost = total_cost / total_orders if total_orders > 0 else 0
            st.metric(
                label="Total Cost (EUR)",
                value=f"€{total_cost:,.2f}",
                delta=f"Avg: €{avg_cost:,.2f}"
            )
        
        with col3:
            total_net = filtered_df['NET_EUR'].sum()
            difference = total_net - total_cost
            diff_color = "normal" if difference >= 0 else "inverse"
            st.metric(
                label="Total NET (EUR)",
                value=f"€{total_net:,.2f}",
                delta=f"Diff: €{difference:,.2f}",
                delta_color=diff_color
            )
        
        with col4:
            unique_accounts = filtered_df['ACCT'].nunique()
            active_accounts = filtered_df[filtered_df['Total cost_EUR'] > 0]['ACCT'].nunique()
            st.metric(
                label="Unique Accounts",
                value=f"{unique_accounts:,}",
                delta=f"Active: {active_accounts}"
            )

        st.markdown("---")
        
        # NEW SECTION: Negative and Cost-Only Accounts Analysis
        st.subheader("Negative Margin & Cost-Only Accounts Analysis")
        
        # Calculate account summaries
        account_summary = filtered_df.groupby(['ACCT', 'ACCT NM']).agg({
            'Total cost_EUR': 'sum',
            'NET_EUR': 'sum',
            'PU COST_EUR': 'sum',
            'SHIP COST_EUR': 'sum',
            'MAN COST_EUR': 'sum',
            'DEL COST_EUR': 'sum',
            'ORD#': 'count'
        }).reset_index()
        
        account_summary['Difference'] = account_summary['NET_EUR'] - account_summary['Total cost_EUR']
        
        # Find negative accounts (negative difference or zero NET)
        negative_accounts = account_summary[
            (account_summary['Difference'] < 0) | 
            ((account_summary['Total cost_EUR'] > 0) & (account_summary['NET_EUR'] == 0))
        ].copy()
        
        if len(negative_accounts) == 0:
            st.success("No accounts with negative margins or cost-only situations found!")
        else:
            # Summary metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(
                    "Problematic Accounts",
                    f"{len(negative_accounts)}",
                    f"Out of {len(account_summary)} total"
                )
            with col2:
                total_loss = negative_accounts['Difference'].sum()
                st.metric(
                    "Total Loss",
                    f"€{abs(total_loss):,.2f}",
                    delta_color="inverse"
                )
            with col3:
                total_cost_negative = negative_accounts['Total cost_EUR'].sum()
                st.metric(
                    "Total Cost of Negative Accounts",
                    f"€{total_cost_negative:,.2f}"
                )
            
            # Sort by loss (most negative first)
            negative_accounts = negative_accounts.sort_values('Difference')
            
            st.markdown("---")
            
            # Overview stacked bar chart
            st.markdown("### Overview: Cost Structure of All Negative Accounts")
            
            fig_overview = go.Figure()
            
            # Add traces for each cost type
            fig_overview.add_trace(go.Bar(
                name='PU Cost',
                x=negative_accounts['ACCT NM'],
                y=negative_accounts['PU COST_EUR'],
                marker_color='#FF6B6B',
                text=[f'€{v:,.0f}' if v > 0 else '' for v in negative_accounts['PU COST_EUR']],
                textposition='inside'
            ))
            fig_overview.add_trace(go.Bar(
                name='Ship Cost',
                x=negative_accounts['ACCT NM'],
                y=negative_accounts['SHIP COST_EUR'],
                marker_color='#4ECDC4',
                text=[f'€{v:,.0f}' if v > 0 else '' for v in negative_accounts['SHIP COST_EUR']],
                textposition='inside'
            ))
            fig_overview.add_trace(go.Bar(
                name='Man Cost',
                x=negative_accounts['ACCT NM'],
                y=negative_accounts['MAN COST_EUR'],
                marker_color='#45B7D1',
                text=[f'€{v:,.0f}' if v > 0 else '' for v in negative_accounts['MAN COST_EUR']],
                textposition='inside'
            ))
            fig_overview.add_trace(go.Bar(
                name='Del Cost',
                x=negative_accounts['ACCT NM'],
                y=negative_accounts['DEL COST_EUR'],
                marker_color='#96CEB4',
                text=[f'€{v:,.0f}' if v > 0 else '' for v in negative_accounts['DEL COST_EUR']],
                textposition='inside'
            ))
            
            fig_overview.update_layout(
                barmode='stack',
                xaxis_title='Account Name',
                yaxis_title='Cost (EUR)',
                height=400,
                showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)
            )
            
            st.plotly_chart(fig_overview, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### Individual Account Cost Breakdowns")
            
            # Create individual breakdowns for each account
            # Use columns to show 2 accounts per row
            for i in range(0, len(negative_accounts), 2):
                cols = st.columns(2)
                
                for j, col in enumerate(cols):
                    if i + j < len(negative_accounts):
                        account = negative_accounts.iloc[i + j]
                        
                        with col:
                            # Create a container for each account
                            with st.container():
                                st.markdown(f"#### {account['ACCT NM']}")
                                st.markdown(f"**Account #:** {account['ACCT']}")
                                
                                # Create pie chart for cost breakdown
                                costs = {
                                    'PU Cost': account['PU COST_EUR'],
                                    'Ship Cost': account['SHIP COST_EUR'],
                                    'Man Cost': account['MAN COST_EUR'],
                                    'Del Cost': account['DEL COST_EUR']
                                }
                                # Filter out zero costs for cleaner pie chart
                                costs = {k: v for k, v in costs.items() if v > 0}
                                
                                if costs:
                                    fig_pie = px.pie(
                                        values=list(costs.values()),
                                        names=list(costs.keys()),
                                        color_discrete_sequence=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'],
                                        hole=0.4
                                    )
                                    fig_pie.update_traces(
                                        textposition='inside',
                                        textinfo='percent+label',
                                        textfont_size=10
                                    )
                                    fig_pie.update_layout(
                                        height=250,
                                        margin=dict(t=20, b=20, l=20, r=20),
                                        showlegend=False
                                    )
                                    st.plotly_chart(fig_pie, use_container_width=True)
                                
                                # Show metrics below the chart
                                subcol1, subcol2, subcol3 = st.columns(3)
                                with subcol1:
                                    st.metric("Total Cost", f"€{account['Total cost_EUR']:,.0f}")
                                with subcol2:
                                    st.metric("NET", f"€{account['NET_EUR']:,.0f}")
                                with subcol3:
                                    loss_value = abs(account['Difference'])
                                    st.metric("Loss", f"€{loss_value:,.0f}", delta_color="inverse")
                                
                                # Detailed breakdown
                                st.markdown("**Cost Details:**")
                                cost_details = []
                                if account['PU COST_EUR'] > 0:
                                    pct = (account['PU COST_EUR'] / account['Total cost_EUR'] * 100)
                                    cost_details.append(f"• PU: €{account['PU COST_EUR']:,.0f} ({pct:.1f}%)")
                                if account['SHIP COST_EUR'] > 0:
                                    pct = (account['SHIP COST_EUR'] / account['Total cost_EUR'] * 100)
                                    cost_details.append(f"• Ship: €{account['SHIP COST_EUR']:,.0f} ({pct:.1f}%)")
                                if account['MAN COST_EUR'] > 0:
                                    pct = (account['MAN COST_EUR'] / account['Total cost_EUR'] * 100)
                                    cost_details.append(f"• Man: €{account['MAN COST_EUR']:,.0f} ({pct:.1f}%)")
                                if account['DEL COST_EUR'] > 0:
                                    pct = (account['DEL COST_EUR'] / account['Total cost_EUR'] * 100)
                                    cost_details.append(f"• Del: €{account['DEL COST_EUR']:,.0f} ({pct:.1f}%)")
                                
                                for detail in cost_details:
                                    st.text(detail)
                                
                                st.text(f"Orders: {account['ORD#']}")
                                
                                # Add separator between accounts
                                st.markdown("---")
            
            # Summary table at the end
            st.markdown("### Summary Table")
            display_df = negative_accounts[['ACCT', 'ACCT NM', 'PU COST_EUR', 'SHIP COST_EUR', 
                                           'MAN COST_EUR', 'DEL COST_EUR', 'Total cost_EUR', 
                                           'NET_EUR', 'Difference', 'ORD#']].copy()
            
            # Format columns
            for col in ['PU COST_EUR', 'SHIP COST_EUR', 'MAN COST_EUR', 'DEL COST_EUR', 
                       'Total cost_EUR', 'NET_EUR', 'Difference']:
                display_df[col] = display_df[col].apply(lambda x: f"€{x:,.2f}")
            
            display_df.columns = ['Account', 'Account Name', 'PU Cost', 'Ship Cost', 
                                 'Man Cost', 'Del Cost', 'Total Cost', 'NET', 'Loss', 'Orders']
            
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Rest of the original dashboard continues below...
        # Create two columns for charts
        col1, col2 = st.columns(2)
        
        with col1:
            # Cost Breakdown Pie Chart
            st.subheader("Cost Breakdown by Type")
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
            # Total Cost by Type
            st.subheader("Total Cost by Type (EUR)")
            
            cost_totals = {
                'PU Cost': filtered_df['PU COST_EUR'].sum(),
                'Ship Cost': filtered_df['SHIP COST_EUR'].sum(),
                'Man Cost': filtered_df['MAN COST_EUR'].sum(),
                'Del Cost': filtered_df['DEL COST_EUR'].sum()
            }
            
            sorted_costs = dict(sorted(cost_totals.items(), key=lambda x: x[1], reverse=False))
            
            fig_cost_totals = px.bar(
                x=list(sorted_costs.values()),
                y=list(sorted_costs.keys()),
                orientation='h',
                color=list(sorted_costs.values()),
                color_continuous_scale='Viridis',
                text=[f'€{v:,.0f}' for v in sorted_costs.values()]
            )
            fig_cost_totals.update_traces(textposition='outside')
            fig_cost_totals.update_layout(
                height=400,
                showlegend=False,
                xaxis_title="Total Amount (EUR)",
                yaxis_title="Cost Type",
                xaxis=dict(tickformat=',.0f')
            )
            st.plotly_chart(fig_cost_totals, use_container_width=True)
        
        # Continue with the rest of the dashboard...
        # (Top 10 Accounts, Monthly Trend, etc. - all without emojis)
        
        # Footer
        st.markdown("---")
        st.caption(f"Dashboard generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Data contains {len(df):,} total records")
        
else:
    # Instructions when no file is uploaded
    st.info("Please upload your Excel file to begin the analysis")
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
    """)
