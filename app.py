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
        st.sidebar.markdown("### 📊 Data Overview")
        st.sidebar.text(f"Total Billed Orders: {len(df)}")
        
        # Display currency distribution info
        if 'CURR' in df.columns:
            currency_counts = df['CURR'].value_counts()
            st.sidebar.markdown("### 💱 Currency Distribution")
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
        
        # Key Metrics - Adjusted for better display
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_orders = len(filtered_df)
            st.metric(
                label="📦 Total Billed Orders",
                value=f"{total_orders:,}",
                delta=f"Status: 440-BILLED"
            )
        
        with col2:
            total_cost = filtered_df['Total cost_EUR'].sum()
            avg_cost = total_cost/total_orders if total_orders > 0 else 0
            st.metric(
                label="💰 Total Cost (EUR)",
                value=f"€{total_cost:,.0f}",
                delta=f"Avg: €{avg_cost:,.0f}"
            )
        
        with col3:
            total_net = filtered_df['NET_EUR'].sum()
            difference = total_net - total_cost
            diff_color = "normal" if difference >= 0 else "inverse"
            st.metric(
                label="📈 Total NET (EUR)",
                value=f"€{total_net:,.0f}",
                delta=f"Margin: €{difference:,.0f}",
                delta_color=diff_color
            )
        
        with col4:
            unique_accounts = filtered_df['ACCT'].nunique()
            active_accounts = filtered_df[filtered_df['Total cost_EUR'] > 0]['ACCT'].nunique()
            st.metric(
                label="👥 Unique Accounts",
                value=f"{unique_accounts:,}",
                delta=f"Active: {active_accounts}"
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
            # CHANGED: Total Cost by Type (instead of count)
            st.subheader("📊 Total Cost by Type (EUR)")
            
            # Calculate total costs for each type
            cost_totals = {
                'PU Cost': filtered_df['PU COST_EUR'].sum(),
                'Ship Cost': filtered_df['SHIP COST_EUR'].sum(),
                'Man Cost': filtered_df['MAN COST_EUR'].sum(),
                'Del Cost': filtered_df['DEL COST_EUR'].sum()
            }
            
            # Sort by value descending
            sorted_costs = dict(sorted(cost_totals.items(), key=lambda x: x[1], reverse=True))
            
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
        
        # Top Accounts by Cost - Ensuring descending order
        st.markdown("---")
        st.subheader("🏆 Top 10 Accounts by Total Cost (Descending)")
        
        account_costs = filtered_df.groupby(['ACCT', 'ACCT NM']).agg({
            'Total cost_EUR': 'sum',
            'NET_EUR': 'sum',
            'ORD#': 'count'
        }).reset_index()
        account_costs.columns = ['Account', 'Account Name', 'Total Cost', 'NET', 'Orders']
        account_costs['Difference'] = account_costs['NET'] - account_costs['Total Cost']
        # Explicitly sort in descending order
        account_costs = account_costs.sort_values('Total Cost', ascending=False).head(10)
        
        # Calculate percentage for better insights
        total_sum = account_costs['Total Cost'].sum()
        account_costs['Percentage'] = (account_costs['Total Cost'] / total_sum * 100).round(1)
        
        fig_top = px.bar(
            account_costs,
            x='Total Cost',
            y='Account Name',
            orientation='h',
            color='Total Cost',
            color_continuous_scale='Blues',
            text='Total Cost',
            hover_data=['Percentage', 'Orders', 'Difference']
        )
        fig_top.update_traces(
            texttemplate='€%{text:,.0f}', 
            textposition='inside',
            textfont_size=10,
            insidetextanchor='middle'
        )
        fig_top.update_layout(
            height=500,
            xaxis_title="Total Cost (EUR)",
            yaxis_title="",
            showlegend=False,
            margin=dict(r=120),  # Add right margin
            xaxis=dict(range=[0, account_costs['Total Cost'].max() * 1.2])  # Extend x-axis
        )
        # Reverse y-axis to show highest at top
        fig_top.update_yaxes(autorange='reversed')
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
        
        # EXPANDED: Detailed Account Analysis Section
        st.markdown("---")
        st.subheader("📋 Account Cost Analysis & Differences")
        
        # Calculate account differences
        account_diff = filtered_df.groupby(['ACCT', 'ACCT NM']).agg({
            'Total cost_EUR': 'sum',
            'NET_EUR': 'sum',
            'ORD#': 'count',
            'TOTAL$_EUR': 'sum'
        }).reset_index()
        account_diff.columns = ['Account', 'Account Name', 'Total Cost', 'NET', 'Orders', 'Total Invoiced']
        account_diff['Difference'] = account_diff['NET'] - account_diff['Total Cost']
        account_diff['Margin %'] = ((account_diff['NET'] - account_diff['Total Cost']) / account_diff['Total Cost'] * 100).round(2)
        
        # Create visualizations for the differences
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 Top 10 Accounts by Margin (NET - Cost)")
            top_margins = account_diff.nlargest(10, 'Difference')
            
            fig_margin = px.bar(
                top_margins,
                x='Difference',
                y='Account Name',
                orientation='h',
                color='Difference',
                color_continuous_scale='RdYlGn',
                text='Difference',
                hover_data=['Total Cost', 'NET', 'Margin %']
            )
            fig_margin.update_traces(texttemplate='€%{text:,.0f}', textposition='outside')
            fig_margin.update_layout(
                height=400,
                xaxis_title="Margin (EUR)",
                yaxis_title="",
                showlegend=False
            )
            fig_margin.update_yaxes(autorange='reversed')
            st.plotly_chart(fig_margin, use_container_width=True)
        
        with col2:
            st.markdown("#### 📈 Margin Percentage Distribution")
            # Filter out infinite margins (where cost is 0)
            margin_data = account_diff[account_diff['Total Cost'] > 0].copy()
            
            fig_margin_pct = px.scatter(
                margin_data,
                x='Total Cost',
                y='Margin %',
                size='Orders',
                color='Margin %',
                color_continuous_scale='RdYlGn',
                hover_data=['Account Name', 'NET', 'Orders'],
                title="Margin % vs Total Cost"
            )
            fig_margin_pct.update_layout(
                height=400,
                xaxis_title="Total Cost (EUR)",
                yaxis_title="Margin %",
                xaxis_type="log"  # Log scale for better visualization
            )
            # Add reference line at 0% margin
            fig_margin_pct.add_hline(y=0, line_dash="dash", line_color="gray", annotation_text="Break-even")
            st.plotly_chart(fig_margin_pct, use_container_width=True)
        
        # Additional analysis charts
        col3, col4 = st.columns(2)
        
        with col3:
            st.markdown("#### 🎯 Cost vs NET Correlation")
            fig_correlation = px.scatter(
                account_diff,
                x='Total Cost',
                y='NET',
                size='Orders',
                color='Margin %',
                color_continuous_scale='Viridis',
                hover_data=['Account Name', 'Difference'],
                trendline="ols"  # Add trend line
            )
            fig_correlation.update_layout(
                height=400,
                xaxis_title="Total Cost (EUR)",
                yaxis_title="NET Amount (EUR)"
            )
            st.plotly_chart(fig_correlation, use_container_width=True)
        
        with col4:
            st.markdown("#### 📊 Pareto Analysis - 80/20 Rule")
            # Sort by Total Cost descending
            pareto_data = account_diff.sort_values('Total Cost', ascending=False).reset_index(drop=True)
            pareto_data['Cumulative Cost'] = pareto_data['Total Cost'].cumsum()
            pareto_data['Cumulative %'] = (pareto_data['Cumulative Cost'] / pareto_data['Total Cost'].sum() * 100).round(1)
            
            # Take top 20 for visibility
            pareto_display = pareto_data.head(20)
            
            fig_pareto = go.Figure()
            
            # Bar chart for individual costs
            fig_pareto.add_trace(go.Bar(
                x=pareto_display.index + 1,
                y=pareto_display['Total Cost'],
                name='Total Cost',
                marker_color='lightblue',
                yaxis='y'
            ))
            
            # Line chart for cumulative percentage
            fig_pareto.add_trace(go.Scatter(
                x=pareto_display.index + 1,
                y=pareto_display['Cumulative %'],
                name='Cumulative %',
                mode='lines+markers',
                marker_color='red',
                yaxis='y2',
                line=dict(width=2)
            ))
            
            # Add 80% reference line
            fig_pareto.add_hline(y=80, line_dash="dash", line_color="green", 
                                annotation_text="80%", yref='y2')
            
            fig_pareto.update_layout(
                height=400,
                xaxis_title="Account Rank",
                yaxis=dict(title="Total Cost (EUR)", side='left'),
                yaxis2=dict(title="Cumulative %", overlaying='y', side='right', range=[0, 100]),
                hovermode='x unified',
                showlegend=True,
                legend=dict(x=0.01, y=0.99)
            )
            
            st.plotly_chart(fig_pareto, use_container_width=True)
            
            # Show Pareto insight
            accounts_80 = pareto_data[pareto_data['Cumulative %'] <= 80].shape[0]
            total_accounts = len(pareto_data)
            if total_accounts > 0:
                st.info(f"💡 **Pareto Insight**: {accounts_80} accounts ({accounts_80/total_accounts*100:.1f}%) contribute to 80% of total costs")
        
        # Display the detailed table
        st.markdown("---")
        st.markdown("### 📊 Detailed Account Table")
        
        # Prepare table data
        account_table = account_diff.copy()
        account_table = account_table.sort_values('Difference', ascending=False, key=abs)
        
        # Format the columns for display
        for col in ['Total Cost', 'NET', 'Difference', 'Total Invoiced']:
            account_table[col] = account_table[col].apply(lambda x: f"€{x:,.2f}")
        account_table['Margin %'] = account_table['Margin %'].apply(lambda x: f"{x:.1f}%")
        
        # Rename columns for display
        account_table = account_table.rename(columns={
            'Total Cost': 'Total Cost (EUR)',
            'NET': 'NET (EUR)',
            'Difference': 'Margin (EUR)',
            'Total Invoiced': 'Invoiced (EUR)'
        })
        
        # Display table with conditional formatting
        st.dataframe(
            account_table.head(20),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Account": st.column_config.TextColumn("Account", width="small"),
                "Account Name": st.column_config.TextColumn("Account Name", width="medium"),
                "Orders": st.column_config.NumberColumn("Orders", width="small"),
                "Margin %": st.column_config.TextColumn("Margin %", width="small"),
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
                data=account_table.to_csv(index=False),
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
