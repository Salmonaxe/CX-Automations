#!/usr/bin/env python3
"""
Cisco Opportunities Timeline Viewer (Integrated Web Version)
Author: Daniel Urgell (durgell@cisco.com)

Interactive web-based GUI for viewing renewal and new opportunity timelines with dynamic filtering.
Run with: streamlit run opps_viewer.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.graph_objects as go
import io

# --- CONSTANTS ---
CISCO_FY_QUARTERS = {
    'Q1': (8, 1),   # August 1
    'Q2': (11, 1),  # November 1
    'Q3': (2, 1),   # February 1
    'Q4': (5, 1)    # May 1
}

# Renewals columns
RENEWALS_REQUIRED_COLUMNS = [
    'Account ARR ($000s)', 'Account Name', 'CX Upsell/PMG', 'Close Date', 'Customer Name',
    'Customer Pulse', 'Deal Id', 'Deal Pulse', 'Expected ATR ($000s)', 'Expiration Date',
    'Expiration Quarter', 'Linked/Related', 'Linked/Related Deals', 'Opportunity Name',
    'Opportunity Owner', 'Opportunity Status', 'Prior ATR ($000s)', 'Product Amount (TCV) ($000s)',
    'Renewal Risk', 'Service Amount (TCV) ($000s)', 'Stage', 'Success Priority'
]

# New Opportunities columns
NEW_OPS_REQUIRED_COLUMNS = [
    'Account Name', 'CX Upsell/PMG', 'Close Date', 'Customer Name',
    'Deal Id', 'Expected Amount TCV ($000s)',
    'Linked/Related', 'Linked/Related Deals', 'Opportunity Name',
    'Opportunity Owner', 'Opportunity Status', 'Stage'
]

RENEWALS_COLUMN_ALIASES = {
    'Renewal Risk': 'RenewalLine Risk'
}

NEW_OPS_COLUMN_ALIASES = {
    'Deal Id': 'Deal ID',
}

MIN_CIRCLE_SIZE = 8
MAX_CIRCLE_SIZE = 32
MIN_DAYS_SEPARATION = 3

COLOR_MAP = {
    'green': '#34A853',
    'yellow': '#FBBC05',
    'red': '#EA4335',
    'grey': '#B0B0B0',
    'black': '#000000',
    'blue': '#0000FF',
    'orange': '#FF8C00',
    'purple': '#800080',
    'cyan': '#00FFFF'
}

# Stage to color mapping for New Opportunities (changed cyan to purple)
STAGE_COLOR_MAP = {
    'qualify': 'black',
    'propose': 'blue',
    'technical validation': 'purple',  # Changed from cyan to purple
    'business validation': 'yellow',
    'negotiate': 'orange',
    'closed won': 'green',
    'closed lost': 'red'
}

# Page config
st.set_page_config(
    page_title="Cisco Opportunities Viewer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def generate_fy_quarters():
    """Generate list of fiscal quarters"""
    quarters = []
    for year in range(24, 30):
        for q in ['Q1', 'Q2', 'Q3', 'Q4']:
            quarters.append(f"{q}FY{year}")
    return quarters

def validate_fy_quarter(fy_str):
    """Validate and parse fiscal quarter"""
    if len(fy_str) != 6:
        return None, None
    quarter = fy_str[:2]
    if quarter not in CISCO_FY_QUARTERS:
        return None, None
    if fy_str[2:4] != 'FY':
        return None, None
    if not fy_str[4:].isdigit():
        return None, None
    
    year = int('20' + fy_str[-2:])
    month, day = CISCO_FY_QUARTERS[quarter]
    
    if quarter == 'Q1':
        start = datetime(year - 1, month, day)
        end = datetime(year - 1, 10, 31)
    elif quarter == 'Q2':
        start = datetime(year - 1, month, day)
        end = datetime(year, 1, 31)
    elif quarter == 'Q3':
        start = datetime(year, month, day)
        end = datetime(year, 4, 30)
    elif quarter == 'Q4':
        start = datetime(year, month, day)
        end = datetime(year, 7, 31)
    
    return start, end

def get_pulse_color(pulse_val):
    """Extract color from pulse value (for renewals)"""
    if isinstance(pulse_val, str) and pulse_val.strip().upper() != 'NA':
        parts = pulse_val.split('-')
        if len(parts) == 2:
            color_name = parts[1].strip().lower()
            return COLOR_MAP.get(color_name, COLOR_MAP['black'])
    return COLOR_MAP['black']

def get_stage_color(stage_val):
    """Extract color from stage value (for new opportunities)"""
    if not stage_val or not isinstance(stage_val, str):
        return COLOR_MAP['black']
    
    try:
        stage_name = stage_val.split("-", 1)[1].strip().lower()
    except IndexError:
        return COLOR_MAP['black']
    
    color_name = STAGE_COLOR_MAP.get(stage_name, 'black')
    return COLOR_MAP.get(color_name, COLOR_MAP['black'])

def get_circle_size(value, min_value, max_value):
    """Calculate circle/square size based on value"""
    if max_value == min_value:
        return MIN_CIRCLE_SIZE
    norm = (value - min_value) / (max_value - min_value)
    return MIN_CIRCLE_SIZE + norm * (MAX_CIRCLE_SIZE - MIN_CIRCLE_SIZE)

def load_and_process_renewals(uploaded_file):
    """Load and process Renewals Excel file"""
    try:
        df = pd.read_excel(uploaded_file)
        
        for expected, actual in RENEWALS_COLUMN_ALIASES.items():
            if expected not in df.columns and actual in df.columns:
                df[expected] = df[actual]
        
        missing = [col for col in RENEWALS_REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            st.error(f"Missing required renewals columns: {', '.join(missing)}")
            return None
        
        df['Expiration Date'] = pd.to_datetime(df['Expiration Date'], errors='coerce')
        
        for col in ['Product Amount (TCV) ($000s)', 'Service Amount (TCV) ($000s)', 
                    'Expected ATR ($000s)', 'Prior ATR ($000s)']:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(r'[\$,]', '', regex=True).replace('nan', '0'),
                    errors='coerce'
                ).fillna(0)
        
        df['Data Type'] = 'Renewal'
        return df
    
    except Exception as e:
        st.error(f"Error loading renewals file: {str(e)}")
        return None

def load_and_process_new_ops(uploaded_file):
    """Load and process New Opportunities Excel file"""
    try:
        df = pd.read_excel(uploaded_file)
        
        for expected, actual in NEW_OPS_COLUMN_ALIASES.items():
            if expected not in df.columns and actual in df.columns:
                df[expected] = df[actual]
        
        missing = [col for col in NEW_OPS_REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            st.error(f"Missing required new opportunities columns: {', '.join(missing)}")
            return None
        
        df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
        
        if 'Expected Amount TCV ($000s)' in df.columns:
            df['Expected Amount TCV ($000s)'] = pd.to_numeric(
                df['Expected Amount TCV ($000s)'].astype(str).str.replace(r'[\$,]', '', regex=True).replace('nan', '0'),
                errors='coerce'
            ).fillna(0)
        
        if 'Deal Id' in df.columns:
            df['Deal Id'] = df['Deal Id'].apply(
                lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, (float, int)) and float(x) == int(x) else str(x)
            )
        
        df['Data Type'] = 'New Opportunity'
        return df
    
    except Exception as e:
        st.error(f"Error loading new opportunities file: {str(e)}")
        return None

def filter_renewals_data(df, start_fy, end_fy, min_atr, account, stage_filter, deal_pulse_filter, customer_pulse_filter):
    """Apply filters to renewals data"""
    fy_start, _ = validate_fy_quarter(start_fy)
    _, fy_end = validate_fy_quarter(end_fy)
    
    if fy_start is None or fy_end is None:
        return None, None
    
    filtered = df[(df['Expiration Date'] >= fy_start) & 
                  (df['Expiration Date'] <= fy_end)].copy()
    
    if filtered.empty:
        return None, None
    
    total_atr_by_deal = filtered.groupby('Deal Id')['Expected ATR ($000s)'].sum().to_dict()
    
    if stage_filter and 'All Stages' not in stage_filter:
        filtered = filtered[filtered['Stage'].isin(stage_filter)]
    
    if deal_pulse_filter and 'All' not in deal_pulse_filter:
        filtered = filtered[filtered['Deal Pulse'].isin(deal_pulse_filter)]
    
    if customer_pulse_filter and 'All' not in customer_pulse_filter:
        filtered = filtered[filtered['Customer Pulse'].isin(customer_pulse_filter)]
    
    if min_atr > 0:
        deal_atr_sum = filtered.groupby('Deal Id')['Expected ATR ($000s)'].sum()
        valid_deals = deal_atr_sum[deal_atr_sum >= min_atr].index
        filtered = filtered[filtered['Deal Id'].isin(valid_deals)]
    
    if account != 'All Accounts':
        filtered = filtered[filtered['Account Name'] == account]
    
    if filtered.empty:
        return None, None
    
    return filtered, total_atr_by_deal

def filter_new_ops_data(df, start_fy, end_fy, min_tcv, account, stage_filter):
    """Apply filters to new opportunities data"""
    fy_start, _ = validate_fy_quarter(start_fy)
    _, fy_end = validate_fy_quarter(end_fy)
    
    if fy_start is None or fy_end is None:
        return None, None
    
    filtered = df[(df['Close Date'] >= fy_start) & 
                  (df['Close Date'] <= fy_end)].copy()
    
    if filtered.empty:
        return None, None
    
    total_tcv_by_deal = filtered.groupby('Deal Id')['Expected Amount TCV ($000s)'].sum().to_dict()
    
    if stage_filter and 'All Stages' not in stage_filter:
        filtered = filtered[filtered['Stage'].isin(stage_filter)]
    
    if min_tcv > 0:
        deal_tcv_sum = filtered.groupby('Deal Id')['Expected Amount TCV ($000s)'].sum()
        valid_deals = deal_tcv_sum[deal_tcv_sum >= min_tcv].index
        filtered = filtered[filtered['Deal Id'].isin(valid_deals)]
    
    if account != 'All Accounts':
        filtered = filtered[filtered['Account Name'] == account]
    
    if filtered.empty:
        return None, None
    
    return filtered, total_tcv_by_deal

def create_integrated_timeline(renewals_df, new_ops_df, fy_start, fy_end, 
                               renewals_totals, new_ops_totals, show_renewals_product, 
                               show_renewals_service, show_new_ops):
    """Create integrated timeline with both renewals and new opportunities"""
    fig = go.Figure()
    
    all_deals = []
    
    # Process Renewals - Product
    if show_renewals_product and renewals_df is not None and not renewals_df.empty:
        renewals_product = renewals_df[renewals_df['Product Amount (TCV) ($000s)'] > 0].copy()
        if not renewals_product.empty:
            agg_renewals_product = (
                renewals_product
                .groupby(['Deal Id', 'Expiration Date'])
                .agg({
                    'Expected ATR ($000s)': 'sum',
                    'Deal Pulse': 'first',
                    'Customer Pulse': 'first',
                    'Account Name': 'first',
                    'Stage': 'first',
                    'Opportunity Name': 'first'
                })
                .reset_index()
            )
            agg_renewals_product['Type'] = 'Renewal-Product'
            agg_renewals_product['Date'] = agg_renewals_product['Expiration Date']
            agg_renewals_product['Value'] = agg_renewals_product['Expected ATR ($000s)']
            all_deals.append(agg_renewals_product)
    
    # Process Renewals - Service
    if show_renewals_service and renewals_df is not None and not renewals_df.empty:
        renewals_service = renewals_df[renewals_df['Service Amount (TCV) ($000s)'] > 0].copy()
        if not renewals_service.empty:
            agg_renewals_service = (
                renewals_service
                .groupby(['Deal Id', 'Expiration Date'])
                .agg({
                    'Expected ATR ($000s)': 'sum',
                    'Deal Pulse': 'first',
                    'Customer Pulse': 'first',
                    'Account Name': 'first',
                    'Stage': 'first',
                    'Opportunity Name': 'first'
                })
                .reset_index()
            )
            agg_renewals_service['Type'] = 'Renewal-Service'
            agg_renewals_service['Date'] = agg_renewals_service['Expiration Date']
            agg_renewals_service['Value'] = agg_renewals_service['Expected ATR ($000s)']
            all_deals.append(agg_renewals_service)
    
    # Process New Opportunities
    if show_new_ops and new_ops_df is not None and not new_ops_df.empty:
        agg_new_ops = (
            new_ops_df
            .groupby(['Deal Id', 'Close Date'])
            .agg({
                'Expected Amount TCV ($000s)': 'sum',
                'Account Name': 'first',
                'Stage': 'first',
                'Opportunity Name': 'first',
                'Opportunity Owner': 'first'
            })
            .reset_index()
        )
        agg_new_ops['Type'] = 'New Opportunity'
        agg_new_ops['Date'] = agg_new_ops['Close Date']
        agg_new_ops['Value'] = agg_new_ops['Expected Amount TCV ($000s)']
        agg_new_ops['Deal Pulse'] = None
        agg_new_ops['Customer Pulse'] = None
        all_deals.append(agg_new_ops)
    
    if not all_deals:
        st.warning("No data to display with current selections")
        return None
    
    # Combine all data
    combined_df = pd.concat(all_deals, ignore_index=True).sort_values('Date')
    
    # Calculate min/max for sizing
    min_value = max(combined_df['Value'].min(), 1)
    max_value = combined_df['Value'].max()
    
    # Assign to timeline rows
    timelines = []
    for idx, row in combined_df.iterrows():
        placed = False
        for timeline in timelines:
            if (row['Date'] - timeline[-1]['Date']).days >= 15:
                timeline.append(row)
                placed = True
                break
        if not placed:
            timelines.append([row])
    
    # Draw timeline lines
    for i in range(len(timelines)):
        fig.add_trace(go.Scatter(
            x=[fy_start, fy_end],
            y=[i, i],
            mode='lines',
            line=dict(color='lightblue', width=2),
            showlegend=False,
            hoverinfo='skip'
        ))
    
    # Plot deals
    for row_idx, timeline in enumerate(timelines):
        for deal in timeline:
            deal_id = str(deal['Deal Id'])
            x = deal['Date']
            y = row_idx
            deal_type = deal['Type']
            stage = str(deal.get('Stage', '')).strip()
            value = float(deal['Value'])
            size = get_circle_size(value, min_value, max_value)
            
            # Determine color and shape
            if deal_type == 'New Opportunity':
                # New Ops: Square, color by stage
                marker_color = get_stage_color(stage)
                marker_symbol = 'square'
                marker_line_color = marker_color
                total_value = new_ops_totals.get(deal['Deal Id'], value) if new_ops_totals else value
                hover_text = (
                    f"[NEW] Deal ID: {deal_id}<br>"
                    f"Account: {deal['Account Name']}<br>"
                    f"TCV: ${int(round(value))}K<br>"
                    f"Total TCV: ${int(round(total_value))}K<br>"
                    f"Stage: {stage}<br>"
                    f"Close Date: {x.strftime('%Y-%m-%d')}<br>"
                    f"Opportunity: {deal['Opportunity Name']}"
                )
                text_color = 'black'
            else:
                # Renewals: Circle, color by Deal Pulse
                deal_pulse_color = get_pulse_color(deal.get('Deal Pulse', 'NA'))
                marker_color = deal_pulse_color
                marker_symbol = 'circle'
                total_value = renewals_totals.get(deal['Deal Id'], value) if renewals_totals else value
                
                is_closed_won = stage.lower() == '6 - closed won'
                is_closed_lost = stage.lower() == '6 - closed lost'
                
                if is_closed_won:
                    marker_color = 'lightgreen'
                    marker_line_color = 'green'
                elif is_closed_lost:
                    marker_color = 'red'
                    marker_line_color = 'darkred'
                else:
                    marker_line_color = marker_color
                
                hover_text = (
                    f"[{deal_type.upper()}] Deal ID: {deal_id}<br>"
                    f"Account: {deal['Account Name']}<br>"
                    f"ATR: ${int(round(value))}K<br>"
                    f"Total ATR: ${int(round(total_value))}K<br>"
                    f"Stage: {stage}<br>"
                    f"Deal Pulse: {deal.get('Deal Pulse', 'NA')}<br>"
                    f"Customer Pulse: {deal.get('Customer Pulse', 'NA')}<br>"
                    f"Expiration: {x.strftime('%Y-%m-%d')}<br>"
                    f"Opportunity: {deal['Opportunity Name']}"
                )
                
                customer_pulse_color = get_pulse_color(deal.get('Customer Pulse', 'NA'))
                text_color = customer_pulse_color if customer_pulse_color != COLOR_MAP['yellow'] else 'black'
            
            # Add marker
            fig.add_trace(go.Scatter(
                x=[x],
                y=[y],
                mode='markers',
                marker=dict(
                    size=size,
                    color=marker_color,
                    symbol=marker_symbol,
                    line=dict(color=marker_line_color, width=2)
                ),
                hovertext=hover_text,
                hoverinfo='text',
                showlegend=False,
                name=deal_type
            ))
            
            # Add text annotation
            is_closed_won = stage.lower() == '6 - closed won'
            is_closed_lost = stage.lower() == '6 - closed lost'
            
            if is_closed_won:
                fig.add_annotation(
                    x=x, y=y, text=deal_id, showarrow=False, yshift=25,
                    font=dict(size=10, color='black', family='Arial Black'),
                    bgcolor='lightgreen', bordercolor='green', borderwidth=2, borderpad=4
                )
            elif is_closed_lost:
                fig.add_annotation(
                    x=x, y=y, text=deal_id, showarrow=False, yshift=25,
                    font=dict(size=10, color='white', family='Arial Black'),
                    bgcolor='red', bordercolor='darkred', borderwidth=2, borderpad=4
                )
            elif deal_type != 'New Opportunity' and text_color == COLOR_MAP['yellow']:
                fig.add_annotation(
                    x=x, y=y, text=deal_id, showarrow=False, yshift=25,
                    font=dict(size=10, color='black', family='Arial Black'),
                    bgcolor='yellow', bordercolor='orange', borderwidth=1, borderpad=4
                )
            else:
                fig.add_annotation(
                    x=x, y=y, text=deal_id, showarrow=False, yshift=25,
                    font=dict(size=10, color=text_color, family='Arial Black')
                )
    
    fig.update_layout(
        title="Integrated Opportunities Timeline",
        xaxis_title="Date",
        yaxis=dict(visible=False),
        height=max(400, len(timelines) * 60 + 100),
        hovermode='closest',
        showlegend=False,
        xaxis=dict(tickformat='%b\n%Y', dtick='M1')
    )
    
    return fig

def create_integrated_details_table(renewals_df, new_ops_df, renewals_totals, new_ops_totals,
                                    show_renewals_product, show_renewals_service, show_new_ops):
    """Create integrated details table"""
    all_data = []
    
    if show_renewals_product and renewals_df is not None and not renewals_df.empty:
        renewals_product = renewals_df[renewals_df['Product Amount (TCV) ($000s)'] > 0].copy()
        if not renewals_product.empty:
            agg = renewals_product.groupby('Deal Id').agg({
                'Account Name': 'first',
                'Expected ATR ($000s)': 'sum',
                'Opportunity Name': 'first',
                'Stage': 'first',
                'Expiration Date': 'first',
                'Deal Pulse': 'first',
                'Customer Pulse': 'first'
            }).reset_index()
            agg['Type'] = 'Renewal-Product'
            agg['Value'] = agg['Expected ATR ($000s)']
            agg['Date'] = agg['Expiration Date']
            agg['Total Value'] = agg['Deal Id'].map(lambda x: renewals_totals.get(x, agg.loc[agg['Deal Id']==x, 'Value'].iloc[0]))
            all_data.append(agg)
    
    if show_renewals_service and renewals_df is not None and not renewals_df.empty:
        renewals_service = renewals_df[renewals_df['Service Amount (TCV) ($000s)'] > 0].copy()
        if not renewals_service.empty:
            agg = renewals_service.groupby('Deal Id').agg({
                'Account Name': 'first',
                'Expected ATR ($000s)': 'sum',
                'Opportunity Name': 'first',
                'Stage': 'first',
                'Expiration Date': 'first',
                'Deal Pulse': 'first',
                'Customer Pulse': 'first'
            }).reset_index()
            agg['Type'] = 'Renewal-Service'
            agg['Value'] = agg['Expected ATR ($000s)']
            agg['Date'] = agg['Expiration Date']
            agg['Total Value'] = agg['Deal Id'].map(lambda x: renewals_totals.get(x, agg.loc[agg['Deal Id']==x, 'Value'].iloc[0]))
            all_data.append(agg)
    
    if show_new_ops and new_ops_df is not None and not new_ops_df.empty:
        agg = new_ops_df.groupby('Deal Id').agg({
            'Account Name': 'first',
            'Expected Amount TCV ($000s)': 'sum',
            'Opportunity Name': 'first',
            'Stage': 'first',
            'Close Date': 'first',
            'Opportunity Owner': 'first'
        }).reset_index()
        agg['Type'] = 'New Opportunity'
        agg['Value'] = agg['Expected Amount TCV ($000s)']
        agg['Date'] = agg['Close Date']
        agg['Total Value'] = agg['Deal Id'].map(lambda x: new_ops_totals.get(x, agg.loc[agg['Deal Id']==x, 'Value'].iloc[0]))
        agg['Deal Pulse'] = 'N/A'
        agg['Customer Pulse'] = 'N/A'
        all_data.append(agg)
    
    if not all_data:
        return None
    
    combined = pd.concat(all_data, ignore_index=True).sort_values('Date')
    
    # Ensure Deal Id is string type for Arrow compatibility
    combined['Deal Id'] = combined['Deal Id'].astype(str)
    
    display_df = pd.DataFrame({
        'Deal Id': combined['Deal Id'].astype(str),  # Explicitly convert to string
        'Type': combined['Type'].astype(str),
        'Account': combined['Account Name'].astype(str),
        'Value': combined['Value'].apply(lambda x: f"${int(round(x))}K"),
        'Total Value': combined['Total Value'].apply(lambda x: f"${int(round(x))}K"),
        'Opportunity': combined['Opportunity Name'].astype(str),
        'Stage': combined['Stage'].astype(str),
        'Date': combined['Date'].dt.strftime('%Y-%m-%d')
    })
    
    return display_df


def display_legend():
    """Display comprehensive legend"""
    st.markdown("""
    ### 🎨 Timeline Legend
    
    **Shapes:**
    - ⭕ **Circle**: Renewal Opportunities
    - ◼️ **Square**: New Opportunities
    
    **Circle Colors (Renewals - Deal Pulse):**
    - 🟢 **Green**: Low risk / Healthy
    - 🟡 **Yellow**: Medium risk / Attention needed
    - 🔴 **Red**: High risk / Critical
    - ⚫ **Grey**: No pulse data
    
    **Square Colors (New Opportunities - Stage):**
    - ⬛ **Black**: Qualify
    - 🟦 **Blue**: Propose
    - 🟪 **Purple**: Technical Validation
    - 🟨 **Yellow**: Business Validation
    - 🟧 **Orange**: Negotiate
    - 🟩 **Green**: Closed Won
    - 🟥 **Red**: Closed Lost
    
    **Size:**
    - Larger shapes = Higher value (ATR for renewals, TCV for new ops)
    
    **Text Frames:**
    - 🟩 **Green Rectangle**: Closed Won
    - 🟥 **Red Rectangle**: Closed Lost
    - 🟨 **Yellow Rectangle**: Customer Pulse Yellow (renewals only)
    
    **Text Color (Renewals only - Customer Pulse):**
    - Deal ID text color reflects Customer Pulse status
    """)


def main():
    st.title("📊 Cisco Opportunities Timeline Viewer (Integrated)")
    st.markdown("---")
    
    with st.sidebar:
        st.header("🎛️ Controls")
        
        st.subheader("📁 File Uploads")
        renewals_file = st.file_uploader("Renewals Excel File", type=['xlsx'], key='renewals')
        new_ops_file = st.file_uploader("New Opportunities Excel File", type=['xlsx'], key='new_ops')
        
        # Load data
        if renewals_file:
            if 'renewals_df' not in st.session_state or st.session_state.get('renewals_file_name') != renewals_file.name:
                with st.spinner("Loading renewals data..."):
                    df = load_and_process_renewals(renewals_file)
                    if df is not None:
                        st.session_state.renewals_df = df
                        st.session_state.renewals_file_name = renewals_file.name
                        st.success(f"✅ Renewals: {len(df)} records")
        
        if new_ops_file:
            if 'new_ops_df' not in st.session_state or st.session_state.get('new_ops_file_name') != new_ops_file.name:
                with st.spinner("Loading new opportunities data..."):
                    df = load_and_process_new_ops(new_ops_file)
                    if df is not None:
                        st.session_state.new_ops_df = df
                        st.session_state.new_ops_file_name = new_ops_file.name
                        st.success(f"✅ New Ops: {len(df)} records")
        
        if 'renewals_df' in st.session_state or 'new_ops_df' in st.session_state:
            st.markdown("---")
            st.subheader("📅 Date Range")
            
            quarters = generate_fy_quarters()
            col1, col2 = st.columns(2)
            with col1:
                start_fy = st.selectbox("Start Quarter", quarters, index=quarters.index('Q3FY26'))
            with col2:
                end_fy = st.selectbox("End Quarter", quarters, index=quarters.index('Q1FY27'))
            
            st.markdown("---")
            st.subheader("📊 Deal Id Selection")
            
            show_renewals_product = st.checkbox("Product Renewals", value=True)
            show_renewals_service = st.checkbox("Service Renewals", value=True)
            show_new_ops = st.checkbox("New Opportunities", value=True)
            
            st.markdown("---")
            st.subheader("🔍 Filters")
            
            min_atr = st.number_input("Min ATR (K) - Renewals", min_value=0, max_value=10000, value=0, step=10)
            min_tcv = st.number_input("Min TCV (K) - New Ops", min_value=0, max_value=10000, value=0, step=10)
            
            # Get combined accounts
            accounts = ['All Accounts']
            if 'renewals_df' in st.session_state:
                accounts.extend(st.session_state.renewals_df['Account Name'].unique().tolist())
            if 'new_ops_df' in st.session_state:
                accounts.extend(st.session_state.new_ops_df['Account Name'].unique().tolist())
            accounts = ['All Accounts'] + sorted(list(set(accounts[1:])))
            
            account = st.selectbox("Account", accounts)
            
            # Get combined stages
            stages = ['All Stages']
            if 'renewals_df' in st.session_state:
                stages.extend(st.session_state.renewals_df['Stage'].dropna().unique().tolist())
            if 'new_ops_df' in st.session_state:
                stages.extend(st.session_state.new_ops_df['Stage'].dropna().unique().tolist())
            stages = ['All Stages'] + sorted(list(set([str(s) for s in stages[1:]])))
            
            stage_filter = st.multiselect("Stage", stages, default=['All Stages'])
            
            # Renewals-specific filters
            if 'renewals_df' in st.session_state:
                st.markdown("**Renewals Filters:**")
                deal_pulses = ['All'] + sorted([str(dp) for dp in st.session_state.renewals_df['Deal Pulse'].dropna().unique()])
                deal_pulse_filter = st.multiselect("Deal Pulse", deal_pulses, default=['All'])
                
                customer_pulses = ['All'] + sorted([str(cp) for cp in st.session_state.renewals_df['Customer Pulse'].dropna().unique()])
                customer_pulse_filter = st.multiselect("Customer Pulse", customer_pulses, default=['All'])
            else:
                deal_pulse_filter = ['All']
                customer_pulse_filter = ['All']
            
            st.markdown("---")
            show_legend = st.checkbox("📖 Show Legend", value=False)
    
    # Main content
    if not renewals_file and not new_ops_file:
        st.info("👈 Please upload at least one Excel file to get started")
        st.markdown("""
        ### Instructions:
        1. Upload Renewals and/or New Opportunities Excel files
        2. Select which data types to display
        3. Adjust filters as needed
        4. View the integrated timeline
        
        ### Features:
        - **Integrated Timeline**: View renewals and new opportunities together
        - **Shape Differentiation**: Circles for renewals, squares for new opportunities
        - **Multiple Filters**: Filter by date, account, stage, pulses, and minimum values
        - **Interactive Visualization**: Hover for details, click to explore
        """)
        
        with st.expander("📖 View Legend", expanded=True):
            display_legend()
        return
    
    if 'renewals_df' not in st.session_state and 'new_ops_df' not in st.session_state:
        st.warning("Please wait while files are being processed...")
        return
    
    if show_legend:
        with st.expander("📖 Legend", expanded=True):
            display_legend()
    
    # Apply filters
    renewals_filtered = None
    renewals_totals = {}
    new_ops_filtered = None
    new_ops_totals = {}
    
    if 'renewals_df' in st.session_state:
        renewals_filtered, renewals_totals = filter_renewals_data(
            st.session_state.renewals_df, start_fy, end_fy, min_atr, account, 
            stage_filter, deal_pulse_filter, customer_pulse_filter
        )
    
    if 'new_ops_df' in st.session_state:
        new_ops_filtered, new_ops_totals = filter_new_ops_data(
            st.session_state.new_ops_df, start_fy, end_fy, min_tcv, account, stage_filter
        )
    
    # Calculate metrics
    total_records = 0
    unique_deals = set()
    total_value = 0
    deals_counted = set()
    
    if renewals_filtered is not None and not renewals_filtered.empty:
        if show_renewals_product:
            prod_df = renewals_filtered[renewals_filtered['Product Amount (TCV) ($000s)'] > 0]
            total_records += len(prod_df)
            for d in prod_df['Deal Id'].unique():
                unique_deals.add(d)
                if d not in deals_counted:
                    deals_counted.add(d)
                    total_value += renewals_totals.get(d, 0)
        if show_renewals_service:
            serv_df = renewals_filtered[renewals_filtered['Service Amount (TCV) ($000s)'] > 0]
            total_records += len(serv_df)
            for d in serv_df['Deal Id'].unique():
                unique_deals.add(d)
                if d not in deals_counted:
                    deals_counted.add(d)
                    total_value += renewals_totals.get(d, 0)
    
    if show_new_ops and new_ops_filtered is not None and not new_ops_filtered.empty:
        total_records += len(new_ops_filtered)
        for d in new_ops_filtered['Deal Id'].unique():
            unique_deals.add(d)
            if d not in deals_counted:
                deals_counted.add(d)
                total_value += new_ops_totals.get(d, 0)
    
    # Display metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Records", total_records)
    with col2:
        st.metric("Unique Deals", len(unique_deals))
    with col3:
        st.metric("Total Value", f"${int(total_value):,}K")
    
    st.markdown("---")
        
    # Timeline
    st.subheader("📈 Integrated Timeline View")
    fy_start, _ = validate_fy_quarter(start_fy)
    _, fy_end = validate_fy_quarter(end_fy)
    
    with st.spinner("Generating timeline..."):
        fig = create_integrated_timeline(
            renewals_filtered, new_ops_filtered, fy_start, fy_end,
            renewals_totals, new_ops_totals,
            show_renewals_product, show_renewals_service, show_new_ops
        )
        if fig:
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
    
    st.markdown("---")
    
    # Details table
    st.subheader("📋 Deal Details")
    details_df = create_integrated_details_table(
        renewals_filtered, new_ops_filtered, renewals_totals, new_ops_totals,
        show_renewals_product, show_renewals_service, show_new_ops
    )
    
    if details_df is not None:
        st.dataframe(details_df, use_container_width=True, height=400, hide_index=True)
        
        csv = details_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Details as CSV",
            data=csv,
            file_name=f"opportunities_{start_fy}_{end_fy}.csv",
            mime="text/csv",
            use_container_width=True
        )

if __name__ == '__main__':
    main()
