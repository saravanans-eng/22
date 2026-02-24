import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from datetime import datetime, timedelta

# Set page configuration
st.set_page_config(
    page_title="Production Status Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 1rem;
    }
    .kpi-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
    }
    .kpi-number {
        font-size: 2.8rem;
        font-weight: 700;
        color: #1E3A8A;
    }
    .kpi-label {
        font-size: 1rem;
        color: #6c757d;
        margin-top: 0.5rem;
    }
    .overdue { color: #dc3545; }
    .held { color: #fd7e14; }
    .pending { color: #ffc107; }
    .today { color: #ffc107; }
    .tomorrow { color: #fd7e14; }
    .advanced { color: #28a745; }
    .metric-box {
        background-color: #f0f2f6;
        border-radius: 5px;
        padding: 10px;
        margin: 5px 0;
    }
    .stDataFrame {
        font-size: 0.9rem;
    }
    .user-card {
        background-color: #262730;
        border-left: 5px solid #1E3A8A;
        padding: 10px;
        margin: 5px 0;
        border-radius: 5px;
    }
    .due-badge {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: 500;
    }
    .due-overdue { background-color: #ffebee; color: #c62828; }
    .due-today { background-color: #fff8e1; color: #f57c00; }
    .due-tomorrow { background-color: #e8f5e9; color: #2e7d32; }
    .due-advanced { background-color: #e3f2fd; color: #1565c0; }
    .stage-card {
        background-color: black;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 10px;
    }
    .stage-name {
        font-size: 1.2rem;
        font-weight: 600;
        color: #1E3A8A;
        border-bottom: 2px solid #1E3A8A;
        padding-bottom: 5px;
        margin-bottom: 10px;
    }
    .stage-count {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 5px 0;
    }
    .count-number {
        font-size: 1.5rem;
        font-weight: 700;
    }
    .count-label {
        font-size: 0.9rem;
        color: #6c757d;
    }
</style>
""", unsafe_allow_html=True)

# Helper functions
@st.cache_data
def load_data(uploaded_file=None):
    """
    Load and preprocess data from the Excel file
    """
    if uploaded_file is not None:
        try:
            # Get sheet names — always specify engine='openpyxl' for .xlsx byte streams
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            # Initialize dataframes
            sheet1_df = None
            sheet2_df = None
            
            # Try to identify sheets by name or position
            for i, sheet_name in enumerate(sheet_names):
                sheet_lower = sheet_name.lower()
                if 'sheet1' in sheet_lower or i == 0:
                    sheet1_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
                elif 'report' in sheet_lower or i == 1:
                    sheet2_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
            
            # If sheets not found by name, use by position
            if sheet1_df is None and len(sheet_names) > 0:
                sheet1_df = pd.read_excel(uploaded_file, sheet_name=sheet_names[0], engine='openpyxl')
            if sheet2_df is None and len(sheet_names) > 1:
                sheet2_df = pd.read_excel(uploaded_file, sheet_name=sheet_names[1], engine='openpyxl')
            
            # Clean and process sheet1 (for inflow data only)
            sheet1_clean = process_sheet1(sheet1_df)
            
            # Clean and process sheet2 (main detailed data)
            sheet2_clean = process_sheet2(sheet2_df)
            
            return sheet1_clean, sheet2_clean
            
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")
            return None, None
    
    return None, None

def process_sheet1(df):
    """
    Process the summary sheet (Sheet1) - Focus on inflow only
    """
    if df is None or df.empty:
        return {'inflow': pd.DataFrame()}
    
    # Extract inflow data (dates with Grand Total)
    inflow_data = []
    
    # Look for date rows
    for idx, row in df.iterrows():
        # Check if first column is a date
        first_val = row.iloc[0] if len(row) > 0 else None
        if pd.notna(first_val):
            try:
                # Try to convert to datetime
                date_val = pd.to_datetime(first_val, errors='coerce')
                if pd.notna(date_val):
                    # Check for Grand Total (column E/index 4) - this is the inflow
                    if len(row) > 4 and pd.notna(row.iloc[4]):
                        inflow_data.append({
                            'Date': date_val,
                            'Inflow': row.iloc[4],
                            'Assigned': row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else 0,
                            'Available': row.iloc[2] if len(row) > 2 and pd.notna(row.iloc[2]) else 0,
                            'Held': row.iloc[3] if len(row) > 3 and pd.notna(row.iloc[3]) else 0
                        })
            except:
                continue
    
    inflow_df = pd.DataFrame(inflow_data)
    
    return {'inflow': inflow_df}

def process_sheet2(df):
    """
    Process the detailed report sheet (report name)
    Focus on Remaining Days column for due calculations
    """
    if df is None or df.empty:
        return pd.DataFrame()
    
    # Clean column names - strip and convert to string
    df.columns = [str(col).strip() for col in df.columns]
    
    # Create a mapping of possible column names
    # Keys use underscores to match all downstream column checks
    column_mapping = {
        'Actual_Date': ['Actual', 'Actual Date', 'actual', 'actual_date'],
        'Journal': ['Journal', 'journal', 'JOURNAL'],
        'Article': ['Article/Vol/Iss', 'Article', 'article', 'vol_iss'],
        'Team': ['Team', 'team', 'TEAM'],
        'Department': ['Department', 'dept', 'DEPT'],
        'Stage': ['Stage', 'stage', 'STAGE'],
        'Process': ['Process', 'process', 'PROCESS'],
        'Assigned_To': ['Assigned To', 'Assigned', 'assigned_to'],
        'Status': ['Status', 'status', 'STATUS'],
        'Task_Status': ['Task Status', 'Task_Status', 'task_status'],
        'Held_Reason': ['Held Reason', 'Held_Reason', 'held_reason'],
        'Remaining_Days': ['Remaining Days', 'Remaining_Days', 'remaining_days'],
        'Assigned_Date': ['Assigned Date', 'Assigned_Date', 'assigned_date'],
        'Due_Date': ['Due Date', 'Due_Date', 'due_date'],
        'Available_Date': ['Available Date', 'Available_Date', 'available_date']
    }
    
    # Find actual column names in the dataframe
    actual_columns = {}
    for std_col, possible_names in column_mapping.items():
        for col in df.columns:
            if any(name.lower() == col.lower() for name in possible_names):
                actual_columns[std_col] = col
                break
    
    # Create result dataframe with found columns
    result_data = {}
    for std_col, actual_col in actual_columns.items():
        result_data[std_col] = df[actual_col]
    
    if not result_data:
        # If no columns found, try to use by position
        col_count = len(df.columns)
        
        if col_count > 2:
            result_data = {
                'Actual_Date': df.iloc[:, 2] if col_count > 2 else None,
                'Journal': df.iloc[:, 3] if col_count > 3 else None,
                'Article': df.iloc[:, 4] if col_count > 4 else None,
                'Team': df.iloc[:, 7] if col_count > 7 else None,
                'Department': df.iloc[:, 8] if col_count > 8 else None,
                'Stage': df.iloc[:, 9] if col_count > 9 else None,
                'Process': df.iloc[:, 11] if col_count > 11 else None,
                'Assigned_To': df.iloc[:, 12] if col_count > 12 else None,
                'Status': df.iloc[:, 14] if col_count > 14 else None,
                'Held_Reason': df.iloc[:, 16] if col_count > 16 else None,
                'Task_Status': df.iloc[:, 22] if col_count > 22 else None,
                'Due_Date': df.iloc[:, 0] if col_count > 0 else None,
                'Assigned_Date': df.iloc[:, 13] if col_count > 13 else None,
                'Available_Date': df.iloc[:, 17] if col_count > 17 else None,
                'Remaining_Days': df.iloc[:, 5] if col_count > 5 else None  # Column F is Remaining Days
            }
    
    # Create dataframe
    result_df = pd.DataFrame(result_data)
    
    if result_df.empty:
        return pd.DataFrame()
    
    # Convert date columns
    date_columns = ['Actual_Date', 'Due_Date', 'Assigned_Date', 'Available_Date']
    for col in date_columns:
        if col in result_df.columns:
            result_df[col] = pd.to_datetime(result_df[col], errors='coerce')
    
    # Use today's date from the file (Feb 23, 2026)
    #today_date = datetime(2026, 2, 23)
    today_date = pd.Timestamp.today().normalize()
    result_df['Today'] = today_date
    
    # CRITICAL: Always recalculate Remaining_Days from Due_Date using today's real date
    # so that due categories reflect the current day, not the stale value stored in Excel.
    if 'Due_Date' in result_df.columns and result_df['Due_Date'].notna().any():
        # Recalculate live from Due_Date column
        result_df['Remaining_Days'] = (result_df['Due_Date'] - result_df['Today']).dt.days
        result_df['Remaining_Days'] = result_df['Remaining_Days'].fillna(0).astype(int)
    elif 'Remaining_Days' in result_df.columns:
        # Fall back to the Excel value if Due_Date is absent, but still sanitise it
        result_df['Remaining_Days'] = pd.to_numeric(result_df['Remaining_Days'], errors='coerce').fillna(0).astype(int)
    else:
        result_df['Remaining_Days'] = 0

    # Create due categories based on Remaining_Days:
    # Overdue: negative values (< 0)
    # Due Today: 0
    # Due Tomorrow: 1
    # Advanced: 2 or more (>= 2)
    conditions = [
        result_df['Remaining_Days'] < 0,                    # Overdue
        result_df['Remaining_Days'] == 0,                   # Due Today
        result_df['Remaining_Days'] == 1,                   # Due Tomorrow
        result_df['Remaining_Days'] >= 2                    # Due in 2 Days or more
    ]
    choices = ['Overdue', 'Today', 'Tomorrow', 'Advanced']

    result_df['Due_Category'] = np.select(conditions, choices, default='Future')

    # Simplified status for easy filtering
    result_df['Due_Status'] = result_df['Due_Category']

    if result_df['Due_Category'].eq('Future').all() and result_df['Remaining_Days'].eq(0).all():
        result_df['Due_Category'] = 'Unknown'
        result_df['Due_Status'] = 'Unknown'
    
    # Calculate Days Since Available for inflow tracking
    if 'Available_Date' in result_df.columns:
        result_df['Days_Since_Available'] = (result_df['Today'] - result_df['Available_Date']).dt.days
        result_df['Available_Category'] = pd.cut(
            result_df['Days_Since_Available'],
            bins=[-float('inf'), 0, 1, 2, 3, 7, float('inf')],
            labels=['Today', '1 Day Ago', '2 Days Ago', '3 Days Ago', 'Within Week', 'Over Week']
        )
    else:
        result_df['Days_Since_Available'] = 0
        result_df['Available_Category'] = 'Unknown'
    
    # Calculate Held Article aging based on assigned date
    if 'Assigned_Date' in result_df.columns and 'Remaining_Days' in result_df.columns:
        result_df['Held_Days'] = (result_df['Today'] - result_df['Assigned_Date']).dt.days
        result_df['Held_Status'] = np.where(
            (result_df['Status'] == 'Held') & (result_df['Remaining_Days'] < 0),
            'Held Overdue',
            np.where(
                (result_df['Status'] == 'Held') & (result_df['Remaining_Days'] >= 0),
                'Held OnDue',
                'Not Held'
            )
        )
    else:
        result_df['Held_Days'] = 0
        result_df['Held_Status'] = 'Unknown'
    
    # Fill NA values
    # IMPORTANT: datetime columns must keep NaT (not be filled with strings)
    # so that .dt accessor continues to work downstream
    for col in result_df.columns:
        try:
            if hasattr(result_df[col], 'dt') and str(result_df[col].dtype) == 'datetime64[ns]':
                pass  # Keep NaT for datetime columns
            elif result_df[col].dtype == 'object':
                result_df[col] = result_df[col].fillna('Unknown')
            else:
                result_df[col] = result_df[col].fillna(0)
        except Exception:
            pass
    
    # Clean up Status values
    if 'Status' in result_df.columns:
        result_df['Status'] = result_df['Status'].astype(str).str.strip()
    else:
        result_df['Status'] = 'Unknown'
    
    if 'Task_Status' in result_df.columns:
        result_df['Task_Status'] = result_df['Task_Status'].astype(str).str.strip()
        result_df['Task_Status'] = result_df['Task_Status'].replace(['nan', 'None', ''], 'OnDue')
    else:
        result_df['Task_Status'] = 'OnDue'
    
    if 'Stage' in result_df.columns:
        result_df['Stage'] = result_df['Stage'].astype(str).str.strip()
        result_df['Stage'] = result_df['Stage'].replace(['nan', 'None', ''], 'Unknown')
    else:
        result_df['Stage'] = 'Unknown'
    
    if 'Process' in result_df.columns:
        result_df['Process'] = result_df['Process'].astype(str).str.strip()
        result_df['Process'] = result_df['Process'].replace(['nan', 'None', ''], 'Unknown')
    else:
        result_df['Process'] = 'Unknown'
    
    if 'Team' in result_df.columns:
        result_df['Team'] = result_df['Team'].astype(str).str.strip()
        result_df['Team'] = result_df['Team'].replace(['nan', 'None', ''], 'Unassigned')
    else:
        result_df['Team'] = 'Unassigned'
    
    if 'Assigned_To' in result_df.columns:
        result_df['Assigned_To'] = result_df['Assigned_To'].astype(str).str.strip()
        result_df['Assigned_To'] = result_df['Assigned_To'].replace(['nan', 'None', ''], 'Unassigned')
        
        # Extract just the name without employee ID if present
        result_df['Assigned_Name'] = result_df['Assigned_To'].apply(
            lambda x: x.split('(')[0].strip() if '(' in x else x
        )
    else:
        result_df['Assigned_To'] = 'Unassigned'
        result_df['Assigned_Name'] = 'Unassigned'
    
    if 'Held_Reason' in result_df.columns:
        result_df['Held_Reason'] = result_df['Held_Reason'].astype(str).str.strip()
        result_df['Held_Reason'] = result_df['Held_Reason'].replace(['nan', 'None', '', '-'], 'No Reason')
    else:
        result_df['Held_Reason'] = 'No Reason'
    
    return result_df

def create_kpi_metrics(sheet2_df):
    """
    Calculate KPI metrics from the detailed data
    Based on Remaining Days column
    """
    if sheet2_df.empty:
        return {
            'total_active': 0,
            'overdue': 0,
            'due_today': 0,
            'due_tomorrow': 0,
            'due_advanced': 0,
            'held': 0,
            'held_overdue': 0,
            'held_ondue': 0,
            'new_inflow_today': 0
        }
    
    # Filter out completed tasks
    active_df = sheet2_df[~sheet2_df['Status'].str.contains('Completed|Complete', case=False, na=False)]
    
    # Due categories based on Remaining Days as per requirements
    overdue = len(active_df[active_df['Remaining_Days'] < 0])
    due_today = len(active_df[active_df['Remaining_Days'] == 0])
    due_tomorrow = len(active_df[active_df['Remaining_Days'] == 1])
    due_advanced = len(active_df[active_df['Remaining_Days'] >= 2])
    
    # Total active tasks
    total_active = len(active_df)
    
    # Held tasks
    held = len(sheet2_df[sheet2_df['Status'] == 'Held'])
    
    # Held Overdue and Held OnDue
    held_overdue = len(sheet2_df[
        (sheet2_df['Status'] == 'Held') & 
        (sheet2_df['Remaining_Days'] < 0)
    ])
    held_ondue = len(sheet2_df[
        (sheet2_df['Status'] == 'Held') & 
        (sheet2_df['Remaining_Days'] >= 0)
    ])
    
    # New inflow today (based on Available Date) - dynamic today
    if 'Available_Date' in sheet2_df.columns:
        today_ts = pd.Timestamp.today().normalize()
        try:
            new_inflow_today = len(sheet2_df[
                (sheet2_df['Available_Date'].dt.normalize() == today_ts) &
                (sheet2_df['Status'] == 'Available')
            ])
        except Exception:
            new_inflow_today = 0
    else:
        new_inflow_today = 0
    
    return {
        'total_active': total_active,
        'overdue': overdue,
        'due_today': due_today,
        'due_tomorrow': due_tomorrow,
        'due_advanced': due_advanced,
        'held': held,
        'held_overdue': held_overdue,
        'held_ondue': held_ondue,
        'new_inflow_today': new_inflow_today
    }

def get_stage_wise_counts(sheet2_df, due_category=None):
    """
    Get stage-wise article counts based on due category
    Returns a dictionary of stage: count
    """
    if sheet2_df.empty:
        return {}
    
    # Filter out completed tasks
    active_df = sheet2_df[~sheet2_df['Status'].str.contains('Completed|Complete', case=False, na=False)]
    
    if due_category:
        # Apply due category filter
        if due_category == 'Overdue':
            filtered_df = active_df[active_df['Remaining_Days'] < 0]
        elif due_category == 'Today':
            filtered_df = active_df[active_df['Remaining_Days'] == 0]
        elif due_category == 'Tomorrow':
            filtered_df = active_df[active_df['Remaining_Days'] == 1]
        elif due_category == 'Advanced':
            filtered_df = active_df[active_df['Remaining_Days'] >= 2]
        else:
            filtered_df = active_df
    else:
        filtered_df = active_df
    
    # Get stage-wise counts
    stage_counts = filtered_df['Stage'].value_counts().to_dict()
    
    # Ensure all stages are represented
    all_stages = sorted(active_df['Stage'].unique())
    result = {}
    for stage in all_stages:
        result[stage] = stage_counts.get(stage, 0)
    
    return result

def create_stage_wise_display(sheet2_df):
    """
    Create the main display showing stage-wise counts for each due category
    This fulfills the requirement: show article count based on stage like:
    Overdue
    S5 - 0
    S100 - 10
    S300 - 0
    """
    if sheet2_df.empty:
        return None, None, None, None
    
    # Get counts for each due category
    overdue_counts = get_stage_wise_counts(sheet2_df, 'Overdue')
    today_counts = get_stage_wise_counts(sheet2_df, 'Today')
    tomorrow_counts = get_stage_wise_counts(sheet2_df, 'Tomorrow')
    advanced_counts = get_stage_wise_counts(sheet2_df, 'Advanced')
    
    return overdue_counts, today_counts, tomorrow_counts, advanced_counts

def create_stage_wise_dataframe(sheet2_df):
    """
    Create a dataframe showing stage-wise counts for all due categories
    """
    if sheet2_df.empty:
        return pd.DataFrame()
    
    # Filter out completed tasks
    active_df = sheet2_df[~sheet2_df['Status'].str.contains('Completed|Complete', case=False, na=False)]
    
    # Get all stages
    stages = sorted(active_df['Stage'].unique())
    
    # Create dataframe
    data = []
    for stage in stages:
        stage_df = active_df[active_df['Stage'] == stage]
        row = {
            'Stage': stage,
            'Overdue': len(stage_df[stage_df['Remaining_Days'] < 0]),
            'Today': len(stage_df[stage_df['Remaining_Days'] == 0]),
            'Tomorrow': len(stage_df[stage_df['Remaining_Days'] == 1]),
            'Advanced (>=2)': len(stage_df[stage_df['Remaining_Days'] >= 2]),
            'Total': len(stage_df)
        }
        data.append(row)
    
    df = pd.DataFrame(data)
    
    # Add total row
    total_row = pd.DataFrame({
        'Stage': ['TOTAL'],
        'Overdue': [df['Overdue'].sum()],
        'Today': [df['Today'].sum()],
        'Tomorrow': [df['Tomorrow'].sum()],
        'Advanced (>=2)': [df['Advanced (>=2)'].sum()],
        'Total': [df['Total'].sum()]
    })
    
    df = pd.concat([df, total_row], ignore_index=True)
    
    return df

def create_due_category_chart(sheet2_df, category):
    """
    Create a bar chart for a specific due category
    """
    if sheet2_df.empty:
        return None
    
    counts = get_stage_wise_counts(sheet2_df, category)
    
    if not counts:
        return None
    
    df = pd.DataFrame(list(counts.items()), columns=['Stage', 'Count'])
    df = df.sort_values('Count', ascending=False)
    
    colors = {
        'Overdue': '#dc3545',
        'Today': '#ffc107',
        'Tomorrow': '#fd7e14',
        'Advanced': '#28a745'
    }
    
    fig = px.bar(
        df,
        x='Stage',
        y='Count',
        title=f'{category} Articles by Stage',
        color_discrete_sequence=[colors.get(category, '#1E3A8A')],
        text='Count'
    )
    
    fig.update_traces(textposition='outside')
    fig.update_layout(height=300, showlegend=False)
    
    return fig

def create_user_assigned_summary(sheet2_df):
    """
    Create User-Process wise count of assigned articles only
    """
    if sheet2_df.empty:
        return pd.DataFrame()
    
    # Filter to only Assigned status
    assigned_df = sheet2_df[sheet2_df['Status'] == 'Assigned']
    
    if assigned_df.empty:
        return pd.DataFrame()
    
    # Create pivot table for users - only Assigned count
    pivot_df = pd.pivot_table(
        assigned_df,
        values='Journal',
        index=['Assigned_Name', 'Process'],
        aggfunc='count',
        fill_value=0
    ).reset_index()
    
    pivot_df.columns = ['User', 'Process', 'Assigned_Count']
    
    # Add total per user
    user_totals = pivot_df.groupby('User')['Assigned_Count'].sum().reset_index()
    user_totals.columns = ['User', 'Total_Assigned']
    
    # Merge totals back
    pivot_df = pivot_df.merge(user_totals, on='User')
    
    # Sort by total assigned descending, then by user
    pivot_df = pivot_df.sort_values(['Total_Assigned', 'User', 'Assigned_Count'], 
                                     ascending=[False, True, False])
    
    # Add rank
    pivot_df['Rank'] = pivot_df.groupby('User').cumcount() + 1
    
    return pivot_df

def create_user_summary_cards(sheet2_df):
    """
    Create summary cards for each user showing their assigned articles
    """
    if sheet2_df.empty:
        return pd.DataFrame()
    
    # Filter to Assigned status
    assigned_df = sheet2_df[sheet2_df['Status'] == 'Assigned']
    
    if assigned_df.empty:
        return pd.DataFrame()
    
    # Get user-wise summary
    user_summary = assigned_df.groupby('Assigned_Name').agg({
        'Journal': 'count',
        'Process': lambda x: x.nunique(),
        'Stage': lambda x: x.nunique(),
        'Remaining_Days': lambda x: (x < 0).sum()  # Count of overdue for this user
    }).reset_index()
    
    user_summary.columns = ['User', 'Total_Assigned', 'Unique_Processes', 'Unique_Stages', 'Overdue_Count']
    
    # Get top processes for each user
    user_processes = assigned_df.groupby(['Assigned_Name', 'Process']).size().reset_index(name='Count')
    user_processes = user_processes.sort_values(['Assigned_Name', 'Count'], ascending=[True, False])
    
    # Add process list as string
    user_process_summary = user_processes.groupby('Assigned_Name').apply(
        lambda x: ', '.join([f"{row['Process']} ({row['Count']})" for _, row in x.head(3).iterrows()])
    ).reset_index(name='Top_Processes')
    
    user_summary = user_summary.merge(user_process_summary, left_on='User', right_on='Assigned_Name', how='left')
    
    # Sort by total assigned
    user_summary = user_summary.sort_values('Total_Assigned', ascending=False)
    
    return user_summary

def find_max_held_user(sheet2_df):
    """
    Find user with maximum held articles
    """
    if sheet2_df.empty:
        return "No data", 0
    
    held_df = sheet2_df[sheet2_df['Status'] == 'Held']
    if held_df.empty:
        return "No held articles", 0
    
    user_held_counts = held_df['Assigned_Name'].value_counts()
    max_user = user_held_counts.index[0]
    max_count = user_held_counts.iloc[0]
    
    return max_user, max_count

def create_inflow_trend_chart(sheet1_data):
    """
    Create inflow trend chart based on available dates
    """
    inflow_df = sheet1_data.get('inflow', pd.DataFrame())
    
    if inflow_df.empty:
        return None
    
    inflow_df = inflow_df.sort_values('Date').tail(14)  # Last 14 days
    
    # Create figure
    fig = go.Figure()
    
    # Add inflow line
    fig.add_trace(
        go.Scatter(x=inflow_df['Date'], y=inflow_df['Inflow'], 
                   mode='lines+markers', 
                   name='Daily Inflow',
                   line=dict(color='#1E3A8A', width=3),
                   marker=dict(size=10),
                   text=inflow_df['Inflow'].round(0),
                   textposition='top center')
    )
    
    # Add breakdown areas
    fig.add_trace(
        go.Bar(x=inflow_df['Date'], y=inflow_df['Assigned'], 
               name='Assigned', marker_color='#1E3A8A', opacity=0.7)
    )
    
    fig.add_trace(
        go.Bar(x=inflow_df['Date'], y=inflow_df['Available'], 
               name='Available', marker_color='#28a745', opacity=0.7)
    )
    
    fig.add_trace(
        go.Bar(x=inflow_df['Date'], y=inflow_df['Held'], 
               name='Held', marker_color='#fd7e14', opacity=0.7)
    )
    
    # Update layout
    fig.update_layout(
        title_text="Daily Inflow Trend with Status Breakdown",
        barmode='stack',
        hovermode='x unified',
        height=400,
        showlegend=True
    )
    
    fig.update_xaxes(title_text="Date")
    fig.update_yaxes(title_text="Number of Tasks")
    
    return fig

def create_due_distribution_chart(sheet2_df):
    """
    Create pie chart of due status distribution
    """
    if sheet2_df.empty:
        return None
    
    # Filter out completed tasks
    filtered_df = sheet2_df[~sheet2_df['Status'].str.contains('Completed|Complete', case=False, na=False)]
    
    # Create due categories based on requirements
    conditions = [
        filtered_df['Remaining_Days'] < 0,
        filtered_df['Remaining_Days'] == 0,
        filtered_df['Remaining_Days'] == 1,
        filtered_df['Remaining_Days'] >= 2
    ]
    choices = ['Overdue', 'Today', 'Tomorrow', 'Advanced']
    
    filtered_df = filtered_df.copy()
    filtered_df['Display_Category'] = np.select(conditions, choices, default='Future')
    
    due_counts = filtered_df['Display_Category'].value_counts().reset_index()
    due_counts.columns = ['Due_Status', 'Count']
    
    colors = {
        'Overdue': '#dc3545',
        'Today': '#ffc107',
        'Tomorrow': '#fd7e14',
        'Advanced': '#28a745',
        'Future': '#6c757d'
    }
    
    fig = px.pie(
        due_counts, 
        values='Count', 
        names='Due_Status',
        title='Tasks by Due Status',
        color='Due_Status',
        color_discrete_map=colors
    )
    
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(height=350)
    
    return fig

def create_held_analysis(sheet2_df):
    """
    Create analysis of held articles by due status
    """
    if sheet2_df.empty:
        return pd.DataFrame()
    
    held_df = sheet2_df[sheet2_df['Status'] == 'Held']
    
    if held_df.empty:
        return pd.DataFrame()
    
    # Group by Stage and Held Status
    pivot_df = pd.pivot_table(
        held_df,
        values='Journal',
        index=['Stage'],
        columns='Held_Status',
        aggfunc='count',
        fill_value=0
    )
    
    # Ensure columns exist
    for status in ['Held OnDue', 'Held Overdue']:
        if status not in pivot_df.columns:
            pivot_df[status] = 0
    
    # Select only relevant columns
    status_cols = [col for col in ['Held OnDue', 'Held Overdue'] if col in pivot_df.columns]
    pivot_df = pivot_df[status_cols]
    
    # Add total
    pivot_df['Total Held'] = pivot_df.sum(axis=1)
    
    # Sort by total held descending
    pivot_df = pivot_df.sort_values('Total Held', ascending=False)
    
    return pivot_df

def main():
    # Header
    st.markdown('<p class="main-header">📊 Production Status Dashboard - Stage-wise Due Analysis</p>', 
                unsafe_allow_html=True)
    
    # Sidebar for file upload and filters
    with st.sidebar:
        st.header("📁 Data Upload")
        uploaded_file = st.file_uploader(
            "Upload DC_Status_Feb_23.xlsx", 
            type=['xlsx', 'xls']
        )
        
        st.markdown("---")
        st.header("🔍 Filters")
        
        # Initialize session state
        if 'filtered_df' not in st.session_state:
            st.session_state.filtered_df = pd.DataFrame()
        
        # Load data
        if uploaded_file:
            with st.spinner("Loading data..."):
                sheet1_data, sheet2_clean = load_data(uploaded_file)
                
                if sheet2_clean is not None and not sheet2_clean.empty:
                    st.session_state.sheet1_data = sheet1_data
                    st.session_state.sheet2_clean = sheet2_clean
                    st.session_state.filtered_df = sheet2_clean.copy()
                    st.success(f"✅ Loaded {len(sheet2_clean)} records")
                else:
                    st.error("No data found in the file")
        else:
            st.info("Please upload the Excel file")
        
        # Apply filters if data exists
        if not st.session_state.filtered_df.empty:
            # Team filter
            teams = ['All'] + sorted(st.session_state.sheet2_clean['Team'].unique().tolist())
            selected_team = st.selectbox("Select Team", teams)
            
            # Stage filter
            stages = ['All'] + sorted(st.session_state.sheet2_clean['Stage'].unique().tolist())
            selected_stage = st.selectbox("Select Stage", stages)
            
            # Process filter
            processes = ['All'] + sorted(st.session_state.sheet2_clean['Process'].unique().tolist())
            selected_process = st.selectbox("Select Process", processes)
            
            # Status filter
            statuses = ['All'] + sorted(st.session_state.sheet2_clean['Status'].unique().tolist())
            selected_status = st.selectbox("Select Status", statuses)
            
            # Apply filters
            filtered_df = st.session_state.sheet2_clean.copy()
            if selected_team != 'All':
                filtered_df = filtered_df[filtered_df['Team'] == selected_team]
            if selected_stage != 'All':
                filtered_df = filtered_df[filtered_df['Stage'] == selected_stage]
            if selected_process != 'All':
                filtered_df = filtered_df[filtered_df['Process'] == selected_process]
            if selected_status != 'All':
                filtered_df = filtered_df[filtered_df['Status'] == selected_status]
            
            st.session_state.filtered_df = filtered_df
            st.info(f"Showing {len(filtered_df)} records")
    
    # Main dashboard area
    if not st.session_state.filtered_df.empty:
        filtered_df = st.session_state.filtered_df
        sheet1_data = st.session_state.sheet1_data
        
        # Calculate KPIs
        kpis = create_kpi_metrics(filtered_df)
        
        # Find user with maximum held articles
        max_held_user, max_held_count = find_max_held_user(filtered_df)
        
        # KPI Row
        st.subheader("📈 Key Performance Indicators")
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number">{kpis['total_active']}</div>
                <div class="kpi-label">Total Active</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number overdue">{kpis['overdue']}</div>
                <div class="kpi-label">Overdue</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number today">{kpis['due_today']}</div>
                <div class="kpi-label">Due Today</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number tomorrow">{kpis['due_tomorrow']}</div>
                <div class="kpi-label">Due Tomorrow</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col5:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number advanced">{kpis['due_advanced']}</div>
                <div class="kpi-label">Due in 2+ Days</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col6:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number held">{kpis['held']}</div>
                <div class="kpi-label">Held Tasks</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Additional KPI Row
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number pending">{kpis['new_inflow_today']}</div>
                <div class="kpi-label">New Inflow Today</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number held">{kpis['held_overdue']}</div>
                <div class="kpi-label">Held Overdue</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number pending">{kpis['held_ondue']}</div>
                <div class="kpi-label">Held OnDue</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-number">{max_held_count}</div>
                <div class="kpi-label">Max Held: {max_held_user}</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # MAIN REQUIREMENT: Stage-wise article counts by due category
        st.subheader("📊 Stage-wise Article Count by Due Status")
        st.markdown("Based on **Remaining Days** column:")
        st.markdown("""
        - **Overdue**: Remaining Days < 0
        - **Due Today**: Remaining Days = 0
        - **Due Tomorrow**: Remaining Days = 1
        - **Due in 2+ Days**: Remaining Days ≥ 2
        """)
        
        # Get stage-wise counts for display
        overdue_counts, today_counts, tomorrow_counts, advanced_counts = create_stage_wise_display(filtered_df)
        
        if overdue_counts:
            # Create 4 columns for the four due categories
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown("### 🔴 Overdue")
                st.markdown(f"<p style='font-size: 2rem; font-weight: bold; color: #dc3545;'>{kpis['overdue']}</p>", unsafe_allow_html=True)
                st.markdown("---")
                for stage, count in overdue_counts.items():
                    st.markdown(f"**{stage}** - {count}")
            
            with col2:
                st.markdown("### 🟡 Due Today")
                st.markdown(f"<p style='font-size: 2rem; font-weight: bold; color: #ffc107;'>{kpis['due_today']}</p>", unsafe_allow_html=True)
                st.markdown("---")
                for stage, count in today_counts.items():
                    st.markdown(f"**{stage}** - {count}")
            
            with col3:
                st.markdown("### 🟠 Due Tomorrow")
                st.markdown(f"<p style='font-size: 2rem; font-weight: bold; color: #fd7e14;'>{kpis['due_tomorrow']}</p>", unsafe_allow_html=True)
                st.markdown("---")
                for stage, count in tomorrow_counts.items():
                    st.markdown(f"**{stage}** - {count}")
            
            with col4:
                st.markdown("### 🟢 Due in 2+ Days")
                st.markdown(f"<p style='font-size: 2rem; font-weight: bold; color: #28a745;'>{kpis['due_advanced']}</p>", unsafe_allow_html=True)
                st.markdown("---")
                for stage, count in advanced_counts.items():
                    st.markdown(f"**{stage}** - {count}")
        else:
            st.info("No stage data available")
        
        st.markdown("---")
        
        # Summary Table
        st.subheader("📋 Stage-wise Summary Table")
        summary_df = create_stage_wise_dataframe(filtered_df)
        if not summary_df.empty:
            # Style the dataframe
            def highlight_due(val):
                if val > 0:
                    if val == 'Overdue':
                        return 'background-color: #ffebee'
                    elif val == 'Today':
                        return 'background-color: #fff8e1'
                    elif val == 'Tomorrow':
                        return 'background-color: #e8f5e9'
                    elif val == 'Advanced (>=2)':
                        return 'background-color: #e3f2fd'
                return ''
            
            st.dataframe(
                summary_df,
                use_container_width=True,
                column_config={
                    'Overdue': st.column_config.NumberColumn('🔴 Overdue'),
                    'Today': st.column_config.NumberColumn('🟡 Today'),
                    'Tomorrow': st.column_config.NumberColumn('🟠 Tomorrow'),
                    'Advanced (>=2)': st.column_config.NumberColumn('🟢 Advanced'),
                    'Total': st.column_config.NumberColumn('📊 Total')
                }
            )
        
        st.markdown("---")
        
        # Row 1: Inflow Trend and Due Distribution
        col1, col2 = st.columns([2, 1])
        
        with col1:
            inflow_chart = create_inflow_trend_chart(sheet1_data)
            if inflow_chart:
                st.plotly_chart(inflow_chart, use_container_width=True)
            else:
                st.info("📊 Insufficient data for inflow trend chart")
        
        with col2:
            due_chart = create_due_distribution_chart(filtered_df)
            if due_chart:
                st.plotly_chart(due_chart, use_container_width=True)
            else:
                st.info("📊 Insufficient data for due distribution chart")
        
        st.markdown("---")
        
        # Held Analysis
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("⏸️ Held Articles Analysis")
            held_analysis = create_held_analysis(filtered_df)
            if not held_analysis.empty:
                st.dataframe(held_analysis, use_container_width=True)
            else:
                st.info("No held articles found")
        
        with col2:
            st.subheader(f"👤 User with Maximum Held Articles")
            st.markdown(f"""
            <div class="user-card">
                <h3 style="color: #1E3A8A; margin:0;">{max_held_user}</h3>
                <p style="font-size: 2rem; font-weight: bold; margin:0;">{max_held_count}</p>
                <p style="color: #6c757d;">articles on hold</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # User-Process Wise Analysis (Assigned Articles Only)
        st.subheader("👥 User-Process Wise Assigned Articles")
        
        user_summary = create_user_assigned_summary(filtered_df)
        if not user_summary.empty:
            # Display summary cards first
            user_cards = create_user_summary_cards(filtered_df)
            
            if not user_cards.empty:
                st.markdown("#### User Summary")
                cols = st.columns(min(4, len(user_cards)))
                for i, (_, row) in enumerate(user_cards.head(8).iterrows()):
                    with cols[i % 4]:
                        overdue_badge = f"<span style='color: #dc3545;'>🔴 {row['Overdue_Count']} overdue</span>" if row['Overdue_Count'] > 0 else ""
                        st.markdown(f"""
                        <div class="user-card">
                            <h4 style="margin:0;">{row['User']}</h4>
                            <p style="font-size: 1.5rem; font-weight: bold; margin:0;">{row['Total_Assigned']}</p>
                            <p style="font-size: 0.8rem; color: #6c757d;">assigned articles {overdue_badge}</p>
                            <p style="font-size: 0.8rem;">Processes: {row['Unique_Processes']} | Stages: {row['Unique_Stages']}</p>
                            <p style="font-size: 0.8rem; background: #333; padding: 3px; border-radius: 3px;">
                                Top: {row['Top_Processes']}
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
            
            st.markdown("#### Detailed Breakdown")
            st.dataframe(user_summary, use_container_width=True, height=500)
        else:
            st.info("No assigned articles found")
        
        st.markdown("---")
        
        # Detailed Pending Tasks Table
        st.subheader("📋 Detailed Pending Tasks")
        
        # Filter to pending tasks (excluding completed)
        pending_df = filtered_df[~filtered_df['Status'].str.contains('Completed|Complete', case=False, na=False)]
        
        if not pending_df.empty:
            # Select columns for display
            display_cols = []
            for col in ['Journal', 'Article', 'Team', 'Stage', 'Process', 'Status', 
                       'Remaining_Days', 'Due_Status', 'Assigned_Name', 'Held_Reason', 
                       'Held_Status', 'Available_Date']:
                if col in pending_df.columns:
                    display_cols.append(col)
            
            display_df = pending_df[display_cols].copy()
            
            # Format Remaining Days with indicators
            def format_remaining_days(val):
                try:
                    val = float(val)
                    if val < 0:
                        return f'🔴 {int(val)} days overdue'
                    elif val == 0:
                        return '🟡 Due today'
                    elif val == 1:
                        return '🟠 Due tomorrow'
                    elif val >= 2:
                        return f'🟢 Due in {int(val)} days'
                    else:
                        return f'⚪ {int(val)} days'
                except:
                    return str(val)
            
            if 'Remaining_Days' in display_df.columns:
                display_df['Due_Indicator'] = display_df['Remaining_Days'].apply(format_remaining_days)
            
            # Format Available Date (handle NaT gracefully)
            if 'Available_Date' in display_df.columns:
                try:
                    display_df['Available_Date'] = display_df['Available_Date'].dt.strftime('%Y-%m-%d').replace({'NaT': ''})
                except Exception:
                    display_df['Available_Date'] = display_df['Available_Date'].astype(str).replace({'NaT': ''})
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=500
            )
            
            # Export option
            if st.button("📥 Export Pending Tasks to CSV"):
                csv = display_df.to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f"pending_tasks_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
        else:
            st.info("📋 No pending tasks found")
    
    else:
        # Display welcome message when no file is uploaded
        st.info("👈 Please upload the DC_Status_Feb_23.xlsx file to view the dashboard")
        
        # Show expected format
        with st.expander("📋 Expected File Format"):
            st.markdown("""
            The Excel file should contain two sheets:
            
            **Sheet1**: Summary data with:
            - Dates in column A
            - Inflow counts (Grand Total) in column E
            - Status breakdowns in columns B-D
            
            **report name**: Detailed task data with:
            - Task details in columns A through Y
            - **Remaining Days in column F** (critical for due calculations)
            - Status information
            - Team assignments
            - Available dates for inflow tracking
            - Assigned dates for held article aging
            
            **Due Categories based on Remaining Days:**
            - **Overdue**: Remaining Days < 0
            - **Due Today**: Remaining Days = 0
            - **Due Tomorrow**: Remaining Days = 1
            - **Due in 2+ Days**: Remaining Days ≥ 2
            """)

if __name__ == "__main__":
    main()