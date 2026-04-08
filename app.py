import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re
import json

# ========== 简单密码保护（可选） ==========
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        pwd = st.sidebar.text_input("Enter system password", type="password")
        if pwd == "admin123":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.sidebar.error("Incorrect password")
            return False
    return True

if not check_password():
    st.stop()

st.set_page_config(page_title="AI Auto Scheduling System", layout="wide")
st.title("📊 AI Auto Scheduling & Progress Tracking System")
st.caption("Auto-parsed from Epicor BAQ Report | Supports multiple files, operation chain, ETA, task completion, alerts, and auto-calibration")

# ========== CSS for warm beige cards ==========
st.markdown("""
<style>
    .stExpander {
        background-color: #fef9e8 !important;
        border-radius: 16px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
        margin-bottom: 16px !important;
        border: 1px solid #fde6b6 !important;
    }
    .stExpander summary {
        background-color: #fef0d5 !important;
        border-radius: 16px 16px 0 0 !important;
        color: #2d2a24 !important;
        font-weight: 600 !important;
        padding: 12px 16px !important;
        font-size: 1rem !important;
        border-bottom: 1px solid #fde6b6 !important;
    }
    .stExpander summary:hover {
        background-color: #fde6b6 !important;
    }
    .stExpander, .stExpander p, .stExpander span, .stExpander label, 
    .stExpander div, .stExpander .stMarkdown, .stExpander .stMetric label,
    .stExpander .stMetric .stMetricValue, .stExpander .stCaption {
        color: #3a3530 !important;
    }
    .stExpander .stMetric label {
        color: #6b5e4e !important;
        font-size: 0.8rem !important;
    }
    .stExpander .stMetric .stMetricValue {
        color: #2d2a24 !important;
        font-weight: 700 !important;
        font-size: 1.2rem !important;
    }
    .stExpander button {
        color: #3a3530 !important;
        background-color: #fef0d5 !important;
        border: 1px solid #e6d3b0 !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
    }
    .stExpander button:hover {
        background-color: #fde6b6 !important;
        border-color: #d4bc8a !important;
    }
    .stExpander .stNumberInput input {
        background-color: #ffffff !important;
        color: #2d2a24 !important;
        border: 1px solid #e6d3b0 !important;
        border-radius: 8px !important;
    }
    .stExpander .stNumberInput input:focus {
        border-color: #d4bc8a !important;
        box-shadow: 0 0 0 1px #d4bc8a !important;
    }
    .stExpander .stAlert {
        background-color: #fefcf5 !important;
        border-left-color: #d4bc8a !important;
        color: #3a3530 !important;
    }
</style>
""", unsafe_allow_html=True)

# ========== Configuration ==========
DEFAULT_LEAD_TIME = {
    'M-LC-FBR': 0.1,
    'P-DB': 0.05,
    'N-MC': 0.3,
    'P-TU-A': 0.1,
    'D-TAP-A': 0.1,
    'P-PCKLNG': 0.1,
    'ASSY-A': 0.1,
    'P-BF': 0.1,
    'W-CDS-A': 0.1,
    'P-DGR': 0.1,
    'W-LWD': 0.1,
    'M-BD': 0.1,
    'P-GRD': 0.1,
    'P-MK-A': 0.1,
    'P-DMK-A': 0.1,
    'F-INK': 0.1,
    '2-PK-A': 0.1,
    'C-SAW': 0.1,
    'F-NPV1': 7.0,
    'DEFAULT': 1.0
}

if 'lead_time_override' not in st.session_state:
    st.session_state.lead_time_override = {}

def get_lead_time(op):
    if op in st.session_state.lead_time_override:
        return st.session_state.lead_time_override[op]
    return DEFAULT_LEAD_TIME.get(op, DEFAULT_LEAD_TIME['DEFAULT'])

OP_TO_DEPT = {
    'P-DB': 'Deburr',
    'M-LC-FBR': 'Laser Cut',
    'P-MK-A': 'Masking',
    'P-DMK-A': 'Demasking',
    'F-INK': 'Inkjet',
    'N-MC': 'Machining',
    'P-TU-A': 'Touch Up',
    'D-TAP-A': 'Tapping',
    'P-PCKLNG': 'Pickling',
    'F-NPV1': 'Passivation',
    'ASSY-A': 'Assembly A',
    'P-BF': 'Buffing',
    'W-CDS-A': 'CD Stud',
    'P-DGR': 'Degreasing',
    'W-LWD': 'Laser Welding',
    'M-BD': 'Bending',
    'P-GRD': 'Grinding',
    '2-PK-A': 'Packing A',
    'C-SAW': 'Sawing',
    'DEFAULT': 'Unassigned'
}

DEPT_CAPACITY = {
    'Deburr': 4, 'Laser Cut': 5, 'Masking': 2, 'Demasking': 2,
    'Inkjet': 1, 'Machining': 4, 'Touch Up': 2, 'Tapping': 2,
    'Pickling': 1, 'Passivation': 2, 'Assembly A': 3, 'Buffing': 3,
    'CD Stud': 4, 'Degreasing': 2,
    'Cutting': 5, 'Laser Welding': 3, 'Bending': 3,
    'Grinding': 2, 'Deburring': 2, 'Packing A': 3, 'Sawing': 2,
    'Unassigned': 5
}

# ========== Helper functions ==========
@st.cache_data
def load_excel(file):
    """读取单个 Excel 文件，返回 DataFrame"""
    df = pd.read_excel(file, header=5)
    df = df.dropna(how='all')
    if 'Main Part Num' in df.columns:
        df['Main Part Num'] = df['Main Part Num'].ffill()
    else:
        st.error(f"File {file.name} missing 'Main Part Num' column")
        st.stop()
    
    for col in ['JobNum/Asm', 'Nesting Num', 'Exwork Date', 'Subpart Qty',
                'Subpart 2D Rev', 'Subpart KK Rev', 'Mtl 10', 'Subpart Part Image',
                'First Process Plan Date', 'Order Date', 'PO - POLine', 'Order Category']:
        if col not in df.columns:
            df[col] = ''
    return df

def load_multiple_excel(files):
    """合并多个 Excel 文件"""
    dfs = []
    for file in files:
        df = load_excel(file)
        dfs.append(df)
    if dfs:
        combined = pd.concat(dfs, ignore_index=True)
        combined = combined.drop_duplicates()
        return combined
    else:
        return pd.DataFrame()

def extract_step_sequence(row):
    steps = []
    step_col_candidates = [f'Step {i}' for i in range(1, 21)]
    if step_col_candidates[0] not in row.index:
        step_col_candidates = [f'Step{i}' for i in range(1, 21)]
    for col in step_col_candidates:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != '':
            steps.append(row[col])
    return steps

def get_next_operation(current_op, steps):
    if not steps:
        return ''
    if pd.isna(current_op) or current_op == '':
        return steps[0] if steps else ''
    if current_op not in steps:
        return ''
    idx = steps.index(current_op)
    if idx + 1 < len(steps):
        return steps[idx + 1]
    else:
        return 'COMPLETED'

def compute_eta(row, today):
    current_op = row.get('Current Operation')
    steps = row['_steps']
    if not steps:
        return today + timedelta(days=7)
    if pd.isna(current_op) or current_op == '' or current_op not in steps:
        remaining_days = sum(get_lead_time(op) for op in steps)
    else:
        try:
            idx = steps.index(current_op)
        except ValueError:
            idx = -1
        remaining_days = 0
        for op in steps[idx+1:]:
            remaining_days += get_lead_time(op)
    remaining_days = max(remaining_days, 0.5)
    return today + timedelta(days=remaining_days)

def compute_remaining_days(row, today):
    current_op = row.get('Current Operation')
    steps = row['_steps']
    if not steps:
        return 7.0
    if pd.isna(current_op) or current_op == '' or current_op not in steps:
        remaining_days = sum(get_lead_time(op) for op in steps)
    else:
        try:
            idx = steps.index(current_op)
        except ValueError:
            idx = -1
        remaining_days = 0
        for op in steps[idx+1:]:
            remaining_days += get_lead_time(op)
    return max(remaining_days, 0.5)

def get_dept_from_op(op):
    if pd.isna(op) or op == '':
        return 'Unassigned'
    return OP_TO_DEPT.get(op, OP_TO_DEPT['DEFAULT'])

def get_planned_date(row):
    if 'First Process Plan Date' in row and pd.notna(row['First Process Plan Date']) and row['First Process Plan Date'] != '':
        return pd.to_datetime(row['First Process Plan Date'], errors='coerce')
    elif 'Order Date' in row and pd.notna(row['Order Date']) and row['Order Date'] != '':
        return pd.to_datetime(row['Order Date'], errors='coerce')
    else:
        return pd.NaT

def get_job_base(job_num):
    if pd.isna(job_num) or job_num == '':
        return ''
    match = re.match(r'^([^-]+)', str(job_num))
    return match.group(1) if match else str(job_num)

def update_main_part_eta(df, today):
    df['_job_base'] = df['JobNum/Asm'].apply(get_job_base)
    df['_is_main'] = df['JobNum/Asm'].astype(str).str.endswith('-0')
    
    subpart_max_eta = {}
    for job_base in df['_job_base'].unique():
        sub_df = df[(df['_job_base'] == job_base) & (~df['_is_main'])]
        if not sub_df.empty:
            subpart_max_eta[job_base] = sub_df['ETA'].max()
    
    for idx, row in df[df['_is_main']].iterrows():
        job_base = row['_job_base']
        remaining_days = compute_remaining_days(row, today)
        sub_max = subpart_max_eta.get(job_base, pd.NaT)
        if pd.notna(sub_max) and row['Current Operation'] != 'COMPLETED':
            new_eta = sub_max + timedelta(days=remaining_days)
        else:
            new_eta = row['ETA']
        df.at[idx, 'ETA'] = new_eta
        df.at[idx, 'Status'] = '✅ On track' if new_eta >= today else '⚠️ Delayed'
    return df

def update_task_to_next_operation(df, index, today):
    row = df.loc[index]
    steps = row['_steps']
    current_op = row['Current Operation']
    if pd.isna(current_op) or current_op == '':
        return df, False, "No current operation"
    if current_op not in steps:
        return df, False, f"Current operation '{current_op}' not found in step chain"
    current_idx = steps.index(current_op)
    if current_idx + 1 >= len(steps):
        df.at[index, 'Current Operation'] = 'COMPLETED'
        df.at[index, 'Current Dept'] = 'Completed'
        df.at[index, 'Next Operation'] = ''
        df.at[index, 'ETA'] = today
    else:
        next_op = steps[current_idx + 1]
        df.at[index, 'Current Operation'] = next_op
        df.at[index, 'Current Dept'] = get_dept_from_op(next_op)
        df.at[index, 'Next Operation'] = get_next_operation(next_op, steps)
        df.at[index, 'ETA'] = compute_eta(df.loc[index], today)
        df.at[index, '_step_start_time'] = datetime.now()
    df.at[index, 'Status'] = '✅ On track' if df.at[index, 'ETA'] >= today else '⚠️ Delayed'
    
    if 'change_log' not in st.session_state:
        st.session_state.change_log = []
    change_entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "job_num": row['JobNum/Asm'],
        "subpart": row['Subpart Part Num'],
        "old_operation": current_op,
        "new_operation": df.at[index, 'Current Operation'],
        "session_id": st.session_state.get('session_id', 'unknown')
    }
    st.session_state.change_log.append(change_entry)
    
    job_base = get_job_base(df.at[index, 'JobNum/Asm'])
    if job_base:
        job_mask = df['_job_base'] == job_base
        main_mask = job_mask & df['_is_main']
        if main_mask.any():
            main_idx = df[main_mask].index[0]
            sub_mask = job_mask & (~df['_is_main'])
            sub_max = df.loc[sub_mask, 'ETA'].max() if sub_mask.any() else pd.NaT
            main_row = df.loc[main_idx]
            remaining_days = compute_remaining_days(main_row, today)
            if pd.notna(sub_max) and main_row['Current Operation'] != 'COMPLETED':
                new_main_eta = sub_max + timedelta(days=remaining_days)
            else:
                new_main_eta = main_row['ETA']
            df.at[main_idx, 'ETA'] = new_main_eta
            df.at[main_idx, 'Status'] = '✅ On track' if new_main_eta >= today else '⚠️ Delayed'
    return df, True, f"Moved to next operation: {next_op if current_idx+1 < len(steps) else 'COMPLETED'}"

def create_gantt_for_job(df, job_base, today):
    job_df = df[df['_job_base'] == job_base].copy()
    if job_df.empty:
        return None
    
    def extract_suffix(job_num):
        match = re.search(r'-(\d+)$', str(job_num))
        if match:
            return int(match.group(1))
        return 0
    
    job_df['_sort_key'] = job_df['JobNum/Asm'].apply(extract_suffix)
    job_df = job_df.sort_values('_sort_key')
    
    job_df['Planned Date'] = pd.to_datetime(job_df['Planned Date'], errors='coerce')
    job_df['ETA'] = pd.to_datetime(job_df['ETA'], errors='coerce')
    job_df['Start'] = job_df['Planned Date'].fillna(pd.Timestamp(today))
    job_df['Finish'] = job_df['ETA'].fillna(pd.Timestamp(today) + pd.Timedelta(days=1))
    mask = job_df['Finish'] < job_df['Start']
    job_df.loc[mask, 'Finish'] = job_df.loc[mask, 'Start'] + pd.Timedelta(days=0.1)
    
    job_df['Current Operation'] = job_df['Current Operation'].fillna('None')
    remaining_days = (job_df['Finish'] - pd.Timestamp(today)).dt.days.clip(lower=0)
    job_df['Remaining Days'] = remaining_days
    job_df['Status'] = job_df['Status']
    job_df['Dept'] = job_df['Current Dept']
    job_df['Task'] = job_df['JobNum/Asm'].astype(str) + ' - ' + job_df['Subpart Part Num'].astype(str)
    
    fig = px.timeline(
        job_df,
        x_start='Start',
        x_end='Finish',
        y='Task',
        color='Dept',
        hover_data={
            'Current Operation': True,
            'Remaining Days': True,
            'Status': True,
            'Start': True,
            'Finish': True,
            'JobNum/Asm': True,
            'Subpart Part Num': True
        },
        title=f"Gantt Chart for Job {job_base} (All Subparts)",
        labels={'Task': 'Job - Subpart', 'Start': 'Planned Start', 'Finish': 'Est. Finish'}
    )
    
    fig.update_yaxes(
        categoryorder='array',
        categoryarray=job_df['Task'].tolist(),
        autorange='reversed'
    )
    
    from datetime import datetime as dt
    today_dt = dt.combine(today, dt.min.time())
    fig.add_shape(
        type='line',
        x0=today_dt, x1=today_dt,
        y0=0, y1=1,
        line=dict(color='red', dash='dash'),
        xref='x', yref='paper'
    )
    fig.add_annotation(
        x=today_dt, y=1,
        text='Today',
        showarrow=False,
        yshift=10,
        xref='x', yref='paper'
    )
    
    fig.update_layout(
        xaxis=dict(
            side='top',
            tickformat='%b %d',
            title=''
        ),
        height=max(700, len(job_df)*50),
        margin=dict(t=60, b=80, l=10, r=10),
        xaxis_title="",
        yaxis_title="Job - Subpart"
    )
    return fig

# ========== 全局筛选函数（仅 Order Category） ==========
def apply_category_filter(df):
    if 'selected_categories' in st.session_state and st.session_state.selected_categories:
        if 'Order Category' in df.columns:
            return df[df['Order Category'].isin(st.session_state.selected_categories)]
    return df

# ========== Sidebar for calibration and filters ==========
st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Auto-Calibration")
if st.sidebar.button("📥 Export Calibration (JSON)"):
    calib_json = json.dumps(st.session_state.lead_time_override, indent=2)
    st.sidebar.download_button("Download", calib_json, file_name="lead_time_calib.json", mime="application/json")

calib_file = st.sidebar.file_uploader("📂 Load Calibration JSON", type=["json"])
if calib_file:
    calib_data = json.load(calib_file)
    st.session_state.lead_time_override = calib_data
    st.sidebar.success("Calibration loaded! Please re-upload the Excel or refresh.")
    st.rerun()

if st.sidebar.button("🔄 Reset All Calibrations"):
    st.session_state.lead_time_override = {}
    st.sidebar.success("Reset to default LEAD_TIME")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🏷️ Order Category Filter")
category_filter_placeholder = st.sidebar.empty()

st.sidebar.markdown("---")
st.sidebar.subheader("📝 Change Log")
if st.sidebar.button("📥 Export Change Log (JSON)"):
    if 'change_log' in st.session_state and st.session_state.change_log:
        log_json = json.dumps(st.session_state.change_log, indent=2)
        st.sidebar.download_button("Download", log_json, file_name="change_log.json", mime="application/json")
    else:
        st.sidebar.warning("No change log to export.")
if st.sidebar.button("🗑️ Clear Change Log"):
    st.session_state.change_log = []
    st.sidebar.success("Change log cleared.")

# ========== Main interface ==========
uploaded_files = st.sidebar.file_uploader(
    "📁 Upload Excel files exported from Epicor (multiple allowed)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    if 'original_df' not in st.session_state or st.sidebar.button("Reload all files"):
        combined_df = load_multiple_excel(uploaded_files)
        if combined_df.empty:
            st.error("No valid data found in the uploaded files.")
            st.stop()
        df = combined_df
        df['_steps'] = df.apply(extract_step_sequence, axis=1)
        today = datetime.now().date()
        df['ETA'] = df.apply(lambda row: compute_eta(row, today), axis=1)
        df['Current Dept'] = df['Current Operation'].apply(get_dept_from_op)
        df['Next Operation'] = df.apply(lambda row: get_next_operation(row['Current Operation'], row['_steps']), axis=1)
        df['Planned Date'] = df.apply(get_planned_date, axis=1)
        df['_job_base'] = df['JobNum/Asm'].apply(get_job_base)
        df['_is_main'] = df['JobNum/Asm'].astype(str).str.endswith('-0')
        df['Status'] = df['ETA'].apply(lambda x: '✅ On track' if x >= today else '⚠️ Delayed')
        df = update_main_part_eta(df, today)
        if 'Exwork Date' in df.columns:
            df['Exwork Date'] = pd.to_datetime(df['Exwork Date'], errors='coerce')
        df['_step_start_time'] = pd.NaT
        st.session_state['original_df'] = df
        st.session_state['df'] = df.copy()
        file_names = ', '.join([f.name for f in uploaded_files])
        st.session_state['file_name'] = file_names
        if 'Order Category' in df.columns:
            all_categories = df['Order Category'].dropna().unique().tolist()
            default_cats = ['New Awarded', 'New Revision']
            selected = [c for c in default_cats if c in all_categories]
            if not selected:
                selected = all_categories
            st.session_state.selected_categories = selected
        st.rerun()
    else:
        df = st.session_state['df']
    
    if 'Order Category' in df.columns:
        all_categories = sorted(df['Order Category'].dropna().unique())
        if all_categories:
            selected_cats = category_filter_placeholder.multiselect(
                "Select Order Categories",
                options=all_categories,
                default=st.session_state.get('selected_categories', all_categories)
            )
            st.session_state.selected_categories = selected_cats
        else:
            category_filter_placeholder.info("No Order Category data")
    else:
        category_filter_placeholder.warning("Column 'Order Category' not found")
    
    filtered_df = apply_category_filter(df.copy())
    st.sidebar.success(f"✅ Loaded {len(df)} total rows from {len(uploaded_files)} file(s), showing {len(filtered_df)} after filter")
    
    # ========== 11 Tabs ==========
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11 = st.tabs([
        "📋 All Items",
        "🏭 Department Workbench",
        "📈 Capacity Dashboard",
        "🔍 Sales Query",
        "📅 Job Gantt Chart",
        "⚠️ Delayed Alerts",
        "📊 Job Progress Board",
        "⏰ Stuck Alerts",
        "📊 Customer Summary",
        "🛠️ Programmer Board",
        "🛠️ Engineering WB Required"
    ])
    
    with tab1:
        st.subheader("Real-time status of all subparts")
        st.caption("**Status explanation**: ✅ On track = Estimated finish date is today or in the future; ⚠️ Delayed = Estimated finish date has passed.\n\n**Note**: For main parts (JobNum/Asm ending with -0), Est. Finish Date = latest subpart finish date + main part's own remaining days.")
        base_cols = ['Main Part Num', 'Subpart Part Num', 'JobNum/Asm', 'Nesting Num',
                     'Current Operation', 'Next Operation', 'Current Dept', 
                     'Planned Date', 'ETA', 'Status', 'Assigned Eng']
        extra_cols = ['Exwork Date', 'Subpart Qty', 'Subpart 2D Rev', 'Subpart KK Rev', 'Mtl 10', 'PO - POLine', 'Order Category']
        display_cols = [c for c in base_cols + extra_cols if c in filtered_df.columns]
        df_display = filtered_df[display_cols].rename(columns={'ETA': 'Est. Finish Date'}).sort_values('Est. Finish Date')
        st.dataframe(df_display, use_container_width=True, height=500)
        
        with st.expander("🔍 View full operation chain for each subpart"):
            for _, row in filtered_df.iterrows():
                if row['_steps']:
                    steps_str = " → ".join(row['_steps'])
                    st.write(f"**{row['Subpart Part Num']}** (Job: {row['JobNum/Asm']}, Nest: {row['Nesting Num']}) : {steps_str}")
    
    with tab2:
        st.subheader("Department to-do list")
        st.info("💡 **JobNum/Asm format**: `-0` indicates the main part; `-1`, `-2` etc. indicate subparts. Main part's Est. Finish Date = max(subpart ETA) + main part's own remaining days.\n\n📊 **Calibration**: Enter actual hours and click Calibrate to adjust future ETAs.")
        dept_list = sorted(filtered_df['Current Dept'].unique())
        selected_dept = st.selectbox("Select department", dept_list, key="dept_select")
        
        col1, col2 = st.columns(2)
        with col1:
            job_filter = st.text_input("Filter by JobNum/Asm (partial match)", key="job_filter")
        with col2:
            nest_filter = st.text_input("Filter by Nesting Num (partial match)", key="nest_filter")
        
        filtered_dept = filtered_df[filtered_df['Current Dept'] == selected_dept]
        if job_filter:
            filtered_dept = filtered_dept[filtered_dept['JobNum/Asm'].astype(str).str.contains(job_filter, case=False, na=False)]
        if nest_filter:
            filtered_dept = filtered_dept[filtered_dept['Nesting Num'].astype(str).str.contains(nest_filter, case=False, na=False)]
        
        filtered_dept = filtered_dept.sort_values('ETA')
        
        if filtered_dept.empty:
            st.info("No tasks match the filters.")
        else:
            for idx, row in filtered_dept.iterrows():
                steps = row['_steps']
                current_op = row['Current Operation']
                if steps and current_op not in ['COMPLETED', ''] and current_op in steps:
                    current_idx = steps.index(current_op)
                    total_steps = len(steps)
                    remaining_steps = total_steps - current_idx - 1
                elif current_op == 'COMPLETED':
                    current_idx = len(steps) - 1 if steps else 0
                    total_steps = len(steps)
                    remaining_steps = 0
                else:
                    current_idx = -1
                    total_steps = len(steps)
                    remaining_steps = total_steps
                
                step_blocks = []
                for i in range(total_steps):
                    if i < current_idx:
                        step_blocks.append("🟩")
                    elif i == current_idx:
                        step_blocks.append("🔵")
                    else:
                        step_blocks.append("⬜")
                step_display = " ".join(step_blocks)
                
                with st.expander(f"📦 {row['JobNum/Asm']} - {row['Subpart Part Num']}", expanded=False):
                    if row['Status'] == '✅ On track':
                        st.markdown('<span style="color:green; font-weight:bold;">✅ On track</span>', unsafe_allow_html=True)
                    else:
                        st.markdown('<span style="color:red; font-weight:bold;">⚠️ Delayed</span>', unsafe_allow_html=True)
                    
                    col_a, col_b, col_c, col_d = st.columns(4)
                    col_a.metric("🔧 Current Op", row['Current Operation'])
                    col_b.metric("🏭 Dept", row['Current Dept'])
                    eta_str = row['ETA'].strftime('%Y-%m-%d') if pd.notna(row['ETA']) else 'Unknown'
                    col_c.metric("📅 Est. Finish", eta_str)
                    exwork_str = row['Exwork Date'].strftime('%Y-%m-%d') if pd.notna(row.get('Exwork Date')) else '-'
                    col_d.metric("🚚 Exwork", exwork_str)
                    
                    st.markdown(f"**工序步骤**  {step_display}")
                    st.caption(f"进度: {current_idx+1}/{total_steps} 步，剩余 {remaining_steps} 个工序")
                    
                    col_btn, col_cal = st.columns(2)
                    with col_btn:
                        if st.button(f"✅ Complete & Next", key=f"complete_{idx}", use_container_width=True):
                            today = datetime.now().date()
                            updated_df, success, message = update_task_to_next_operation(st.session_state['df'], idx, today)
                            if success:
                                st.session_state['df'] = updated_df
                                st.rerun()
                            else:
                                st.error(f"Failed: {message}")
                    with col_cal:
                        op = row['Current Operation']
                        if op not in ['COMPLETED', '']:
                            actual_hours = st.number_input(f"Actual hrs", min_value=0.0, step=0.5, key=f"actual_{idx}", label_visibility="collapsed", placeholder="Hours")
                            if st.button(f"Calibrate", key=f"calib_{idx}", use_container_width=True):
                                if actual_hours > 0:
                                    old_days = get_lead_time(op)
                                    old_hours = old_days * 10
                                    new_hours = 0.7 * old_hours + 0.3 * actual_hours
                                    new_days = new_hours / 10
                                    st.session_state.lead_time_override[op] = new_days
                                    st.success(f"Calibrated {op}: {old_days:.2f} days → {new_days:.2f} days")
                                    st.rerun()
                                else:
                                    st.warning("Enter actual hours first")
                        else:
                            st.write("✅ Completed")
        
        if st.button("📥 Download updated Excel (with progress changes)"):
            output_df = st.session_state['df'].drop(columns=['_steps', '_job_base', '_is_main', '_step_start_time'], errors='ignore')
            for col in ['ETA', 'Exwork Date', 'Planned Date']:
                if col in output_df.columns:
                    output_df[col] = output_df[col].astype(str)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='UpdatedSchedule')
            st.download_button(
                label="Download Excel",
                data=output.getvalue(),
                file_name=f"updated_{st.session_state['file_name']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with tab3:
        st.subheader("Department capacity load")
        dept_load = filtered_df['Current Dept'].value_counts().reset_index()
        dept_load.columns = ['Department', 'Task Count']
        dept_load['Capacity'] = dept_load['Department'].map(DEPT_CAPACITY).fillna(5)
        dept_load['Load (%)'] = (dept_load['Task Count'] / dept_load['Capacity'] * 100).round(1)
        dept_load = dept_load.sort_values('Load (%)', ascending=False)
        fig = px.bar(dept_load, x='Department', y='Task Count', color='Load (%)',
                     title='Current task load by department (darker = busier)',
                     labels={'Task Count': 'Number of tasks'})
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(dept_load, use_container_width=True)
        overload = dept_load[dept_load['Load (%)'] > 100]
        if not overload.empty:
            st.error(f"⚠️ Overloaded departments: {', '.join(overload['Department'].tolist())}")
    
    with tab4:
        st.subheader("Quick sales query")
        st.info("💡 Enter a Job Number, PO Number, or Subpart Part Num to see summary and sorted subpart list.")
        default_search = st.session_state.pop('selected_job_sales', '')
        search_term = st.text_input("Enter Job Number, PO-POLine, or Subpart Part Num", value=default_search, key="sales_query")
        if search_term:
            mask = (filtered_df['_job_base'].astype(str).str.contains(search_term, case=False, na=False) |
                    filtered_df['Subpart Part Num'].str.contains(search_term, case=False, na=False) |
                    filtered_df['PO - POLine'].astype(str).str.contains(search_term, case=False, na=False))
            result = filtered_df[mask].copy()
            if not result.empty:
                def extract_suffix(job_num):
                    match = re.search(r'-(\d+)$', str(job_num))
                    if match:
                        return int(match.group(1))
                    return 0
                result['_sort_key'] = result['JobNum/Asm'].apply(extract_suffix)
                result = result.sort_values('_sort_key')
                
                overall_eta = result['ETA'].max()
                overall_exwork = result['Exwork Date'].max()
                total_subparts = len(result)
                on_track = len(result[result['Status'] == '✅ On track'])
                delayed = total_subparts - on_track
                dept_counts = result['Current Dept'].value_counts()
                bottleneck_dept = dept_counts.index[0] if not dept_counts.empty else 'None'
                po_jobs = result['_job_base'].unique()
                
                st.markdown("### 📊 Search Summary")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Subparts Found", total_subparts)
                col2.metric("On Track", on_track)
                col3.metric("Delayed", delayed)
                st.info(f"**Overall Est. Finish Date:** {overall_eta.strftime('%Y-%m-%d') if pd.notna(overall_eta) else 'Unknown'}  |  **Overall Exwork Date:** {overall_exwork.strftime('%Y-%m-%d') if pd.notna(overall_exwork) else 'Not set'}  |  **Bottleneck Dept:** {bottleneck_dept}")
                if len(po_jobs) > 1:
                    st.write(f"**Jobs involved in this PO:** {', '.join(po_jobs)}")
                
                st.markdown("### 📋 Subpart Details (sorted by -0, -1, -2...)")
                filter_text = st.text_input("🔍 Filter table (search any column)", key="subpart_filter", placeholder="e.g., Deburr, On track, P-DB...")
                display_cols = ['JobNum/Asm', 'Subpart Part Num', 'Current Operation', 'Current Dept', 
                                'ETA', 'Status', 'Exwork Date', 'Subpart Qty', 'PO - POLine', 'Order Category']
                display_cols = [c for c in display_cols if c in result.columns]
                result_display = result[display_cols].rename(columns={'ETA': 'Est. Finish Date'})
                if filter_text:
                    mask_filter = result_display.astype(str).apply(lambda row: row.str.contains(filter_text, case=False).any(), axis=1)
                    result_display = result_display[mask_filter]
                    if result_display.empty:
                        st.warning("No rows match the filter.")
                st.dataframe(result_display, use_container_width=True)
                
                with st.expander("🔍 View full operation chain for each subpart"):
                    for _, row in result.iterrows():
                        if row['_steps']:
                            steps_str = " → ".join(row['_steps'])
                            st.write(f"**{row['JobNum/Asm']} - {row['Subpart Part Num']}** (PO: {row['PO - POLine']}) : {steps_str}")
            else:
                st.warning("No matching Job, PO, or Subpart found.")
    
    with tab5:
        st.subheader("📅 Job Gantt Chart - Subpart Progress Visualization")
        st.caption("Select a Job to view its Gantt chart. Each bar represents a subpart from Planned Start to Estimated Finish Date. Color indicates current department. The red dashed line marks today.")
        all_jobs = sorted(filtered_df['_job_base'].dropna().unique())
        if len(all_jobs) == 0:
            st.warning("No Job numbers found in the data.")
        else:
            default_job = st.session_state.pop('selected_job_gantt', None)
            if default_job and default_job in all_jobs:
                default_index = all_jobs.index(default_job)
            else:
                default_index = 0
            selected_job = st.selectbox("Select Job Number (Base)", all_jobs, index=default_index)
            fig = create_gantt_for_job(filtered_df, selected_job, datetime.now().date())
            if fig:
                st.plotly_chart(fig, use_container_width=True, key=f"gantt_{selected_job}")
                job_data = filtered_df[filtered_df['_job_base'] == selected_job]
                total_subparts = len(job_data)
                st.metric("Total Subparts", total_subparts)
            else:
                st.error("Failed to generate Gantt chart. Please check data.")
    
    with tab6:
        st.subheader("⚠️ Delayed Tasks Alert Dashboard")
        delayed_df = filtered_df[filtered_df['Status'] == '⚠️ Delayed'].copy()
        if delayed_df.empty:
            st.success("🎉 No delayed tasks! All on track.")
        else:
            st.error(f"🚨 {len(delayed_df)} task(s) are delayed.")
            dept_delay = delayed_df['Current Dept'].value_counts().reset_index()
            dept_delay.columns = ['Department', 'Delayed Count']
            fig_delay = px.bar(dept_delay, x='Department', y='Delayed Count', title='Delayed Tasks by Department', color='Delayed Count')
            st.plotly_chart(fig_delay, use_container_width=True)
            
            st.subheader("Delayed Task List")
            delay_cols = ['JobNum/Asm', 'Subpart Part Num', 'Current Dept', 'Current Operation', 'ETA', 'Planned Date']
            delay_cols = [c for c in delay_cols if c in delayed_df.columns]
            today_date = datetime.now().date()
            delayed_df['Delayed Days'] = (today_date - delayed_df['ETA']).dt.days
            delayed_display = delayed_df[delay_cols + ['Delayed Days']].sort_values('Delayed Days', ascending=False)
            st.dataframe(delayed_display, use_container_width=True)
            
            st.subheader("Job Summary with Delays")
            job_delay = delayed_df.groupby('_job_base').size().reset_index(name='Delayed Subparts')
            st.dataframe(job_delay, use_container_width=True)
    
    with tab7:
        st.subheader("📊 Global Job Progress Board")
        st.caption("Overview of all Jobs: estimated finish dates, progress, bottleneck departments, and more.")
        
        job_group = filtered_df.groupby('_job_base').agg({
            'Subpart Part Num': 'count',
            'ETA': lambda x: x.max(),
            'Status': lambda x: (x == '✅ On track').sum(),
            'Current Dept': lambda x: x.mode()[0] if not x.empty else 'Unknown',
            'Exwork Date': lambda x: x.max(),
            'JobNum/Asm': lambda x: next(iter(x), '')
        }).reset_index()
        job_group.columns = ['Job', 'Subpart Count', 'Main Part ETA', 'On Track Count', 'Bottleneck Dept', 'Exwork Date', 'Sample JobNum']
        job_group['Delayed Count'] = job_group['Subpart Count'] - job_group['On Track Count']
        job_group['Progress %'] = (job_group['On Track Count'] / job_group['Subpart Count'] * 100).round(1).fillna(0)
        
        job_group['Main Part ETA'] = pd.to_datetime(job_group['Main Part ETA'], errors='coerce')
        job_group['Exwork Date'] = pd.to_datetime(job_group['Exwork Date'], errors='coerce')
        job_group = job_group.sort_values('Main Part ETA')
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Jobs", len(job_group))
        col2.metric("Jobs with Delayed Subparts", len(job_group[job_group['Delayed Count'] > 0]))
        col3.metric("Fully On Track Jobs", len(job_group[job_group['Delayed Count'] == 0]))
        
        display_cols = ['Job', 'Sample JobNum', 'Subpart Count', 'Progress %', 'Main Part ETA', 'Delayed Count', 'Bottleneck Dept', 'Exwork Date']
        display_df = job_group[display_cols].copy()
        display_df['Main Part ETA'] = display_df['Main Part ETA'].dt.strftime('%Y-%m-%d')
        display_df['Exwork Date'] = display_df['Exwork Date'].dt.strftime('%Y-%m-%d')
        display_df = display_df.rename(columns={
            'Job': 'Job Base',
            'Sample JobNum': 'JobNum/Asm (Sample)',
            'Subpart Count': 'Subparts',
            'Progress %': 'Progress (%)',
            'Main Part ETA': 'Est. Finish Date',
            'Delayed Count': 'Delayed Subparts',
            'Bottleneck Dept': 'Bottleneck Dept'
        })
        st.dataframe(display_df, use_container_width=True, height=400)
        
        st.subheader("Quick Actions")
        selected_job_for_action = st.selectbox("Select Job to view details", job_group['Job'].tolist())
        col_gantt, col_sales = st.columns(2)
        
        if 'show_gantt' not in st.session_state:
            st.session_state.show_gantt = False
            st.session_state.show_sales = False
            st.session_state.gantt_job = None
            st.session_state.sales_job = None
        
        with col_gantt:
            if st.button("View Gantt Chart for this Job"):
                st.session_state.show_gantt = True
                st.session_state.show_sales = False
                st.session_state.gantt_job = selected_job_for_action
                st.session_state.sales_job = None
                st.rerun()
        with col_sales:
            if st.button("View Sales Summary for this Job"):
                st.session_state.show_sales = True
                st.session_state.show_gantt = False
                st.session_state.sales_job = selected_job_for_action
                st.session_state.gantt_job = None
                st.rerun()
        
        if st.session_state.show_gantt and st.session_state.gantt_job:
            st.markdown("---")
            st.subheader(f"📅 Gantt Chart for Job {st.session_state.gantt_job}")
            fig = create_gantt_for_job(filtered_df, st.session_state.gantt_job, datetime.now().date())
            if fig:
                st.plotly_chart(fig, use_container_width=True, key=f"gantt_{st.session_state.gantt_job}")
            else:
                st.warning("Could not generate Gantt chart.")
        
        if st.session_state.show_sales and st.session_state.sales_job:
            st.markdown("---")
            st.subheader(f"📊 Sales Summary for Job {st.session_state.sales_job}")
            result = filtered_df[filtered_df['_job_base'] == st.session_state.sales_job].copy()
            if not result.empty:
                def extract_suffix(job_num):
                    match = re.search(r'-(\d+)$', str(job_num))
                    if match:
                        return int(match.group(1))
                    return 0
                result['_sort_key'] = result['JobNum/Asm'].apply(extract_suffix)
                result = result.sort_values('_sort_key')
                total_subparts = len(result)
                on_track = len(result[result['Status'] == '✅ On track'])
                delayed = total_subparts - on_track
                main_part_row = result[result['JobNum/Asm'].astype(str).str.endswith('-0')]
                if not main_part_row.empty:
                    main_eta = main_part_row.iloc[0]['ETA'].strftime('%Y-%m-%d') if pd.notna(main_part_row.iloc[0]['ETA']) else 'Unknown'
                else:
                    main_eta = 'No main part'
                exwork_dates = result['Exwork Date'].dropna()
                exwork_date = exwork_dates.max().strftime('%Y-%m-%d') if not exwork_dates.empty else 'Not set'
                dept_counts = result['Current Dept'].value_counts()
                bottleneck_dept = dept_counts.index[0] if not dept_counts.empty else 'None'
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Subparts", total_subparts)
                col2.metric("On Track", on_track)
                col3.metric("Delayed", delayed)
                st.info(f"**Main Part Est. Finish:** {main_eta}  |  **Exwork Date:** {exwork_date}  |  **Bottleneck Dept:** {bottleneck_dept}")
                st.markdown("#### Subpart Details")
                display_cols = ['JobNum/Asm', 'Subpart Part Num', 'Current Operation', 'Current Dept', 
                                'ETA', 'Status', 'Exwork Date', 'Subpart Qty', 'PO - POLine', 'Order Category']
                display_cols = [c for c in display_cols if c in result.columns]
                result_display = result[display_cols].rename(columns={'ETA': 'Est. Finish Date'})
                st.dataframe(result_display, use_container_width=True)
            else:
                st.warning("No data found for this Job.")
    
    with tab8:
        st.subheader("⏰ Stuck Tasks Alert (Exceeding Custom Time Threshold)")
        st.caption("Tasks that have been in the same operation longer than the user-defined threshold (hours). Only tasks that were advanced via 'Complete & Next' are tracked.")
        stuck_hours = st.number_input("Alert when a task stays in the same operation longer than (hours)", min_value=1, max_value=168, value=24, step=1, help="Set threshold in hours. Tasks exceeding this time will appear as stuck.")
        stuck_days = stuck_hours / 24.0
        
        stuck_df = filtered_df[filtered_df['_step_start_time'].notna() & (filtered_df['Current Operation'] != 'COMPLETED')].copy()
        if stuck_df.empty:
            st.success("🎉 No tasks with tracked start time. Advance tasks via 'Complete & Next' to monitor.")
        else:
            now = datetime.now()
            stuck_df['Stayed Days'] = (now - stuck_df['_step_start_time']).dt.total_seconds() / 86400.0
            stuck_df['Threshold Days'] = stuck_days
            stuck_df['Exceed'] = stuck_df['Stayed Days'] > stuck_df['Threshold Days']
            stuck_df['Exceed Ratio'] = (stuck_df['Stayed Days'] / stuck_df['Threshold Days']).round(2)
            stuck_df['Status'] = stuck_df['Exceed'].apply(lambda x: '🔴 Stuck' if x else '🟡 Within threshold')
            
            stuck_only = stuck_df[stuck_df['Exceed'] == True]
            if stuck_only.empty:
                st.success(f"🎉 No tasks exceed the {stuck_hours}-hour threshold.")
            else:
                st.error(f"🚨 {len(stuck_only)} task(s) have exceeded the {stuck_hours}-hour threshold.")
                dept_stuck = stuck_only['Current Dept'].value_counts().reset_index()
                dept_stuck.columns = ['Department', 'Stuck Count']
                fig_stuck = px.bar(dept_stuck, x='Department', y='Stuck Count', title='Stuck Tasks by Department', color='Stuck Count')
                st.plotly_chart(fig_stuck, use_container_width=True)
                
                st.subheader("Stuck Task List")
                display_cols = ['JobNum/Asm', 'Subpart Part Num', 'Current Operation', 'Current Dept', 
                                '_step_start_time', 'Stayed Days', 'Threshold Days', 'Exceed Ratio', 'Status']
                display_cols = [c for c in display_cols if c in stuck_only.columns]
                stuck_display = stuck_only[display_cols].copy()
                stuck_display['_step_start_time'] = stuck_display['_step_start_time'].dt.strftime('%Y-%m-%d %H:%M')
                stuck_display = stuck_display.rename(columns={
                    '_step_start_time': 'Start Time',
                    'Stayed Days': 'Stayed (days)',
                    'Threshold Days': 'Threshold (days)',
                    'Exceed Ratio': 'Ratio'
                })
                st.dataframe(stuck_display, use_container_width=True)
    
    with tab9:
        st.subheader("📊 Customer Summary - Main Parts Only (based on Exwork Date)")
        st.caption("Aggregates main parts (JobNum/Asm ending with -0) per customer and month using Exwork Date. Shows total main parts per customer.")
        
        main_parts_df = filtered_df[filtered_df['JobNum/Asm'].astype(str).str.endswith('-0')].copy()
        if main_parts_df.empty:
            st.warning("No main parts (ending with -0) found in the data.")
        else:
            if 'Exwork Date' not in main_parts_df.columns or main_parts_df['Exwork Date'].isna().all():
                st.error("No valid Exwork Date found for main parts. Please ensure the Excel contains 'Exwork Date' column.")
            else:
                main_parts_df['Customer'] = main_parts_df['Main Part Num'].fillna('Unknown').apply(
                    lambda x: x.split('-')[0] if '-' in x else x
                )
                exwork_date_col = 'Exwork Date'
                main_parts_df[exwork_date_col] = pd.to_datetime(main_parts_df[exwork_date_col], errors='coerce')
                main_parts_df = main_parts_df.dropna(subset=[exwork_date_col])
                
                if main_parts_df.empty:
                    st.warning("No main parts with valid Exwork Date.")
                else:
                    main_parts_df['YearMonth'] = main_parts_df[exwork_date_col].dt.strftime('%Y-%m')
                    monthly_agg = main_parts_df.groupby(['Customer', 'YearMonth']).size().reset_index(name='Main Part Count')
                    pivot = monthly_agg.pivot(index='Customer', columns='YearMonth', values='Main Part Count').fillna(0).astype(int)
                    pivot['Total Main Parts'] = pivot.sum(axis=1)
                    pivot = pivot.sort_values('Total Main Parts', ascending=False)
                    st.dataframe(pivot, use_container_width=True, height=400)
                    
                    st.subheader("📈 Customer Daily Trend (based on Order Date)")
                    customers = sorted(main_parts_df['Customer'].unique())
                    selected_customer = st.selectbox("Select Customer", customers)
                    
                    if selected_customer:
                        cust_data = main_parts_df[main_parts_df['Customer'] == selected_customer].copy()
                        if 'Order Date' not in cust_data.columns or cust_data['Order Date'].isna().all():
                            st.warning("No valid Order Date for this customer. Cannot display daily trend.")
                        else:
                            order_date_col = 'Order Date'
                            cust_data[order_date_col] = pd.to_datetime(cust_data[order_date_col], errors='coerce')
                            cust_data = cust_data.dropna(subset=[order_date_col])
                            cust_data['Date'] = cust_data[order_date_col].dt.date
                            daily_counts = cust_data.groupby('Date').size().reset_index(name='New Main Parts')
                            today_date = datetime.now().date()
                            start_date = today_date - timedelta(days=60)
                            date_range = pd.date_range(start=start_date, end=today_date, freq='D').date
                            daily_counts = daily_counts.set_index('Date').reindex(date_range, fill_value=0).reset_index()
                            daily_counts.columns = ['Date', 'New Main Parts']
                            
                            yesterday = today_date - timedelta(days=1)
                            yesterday_count = daily_counts[daily_counts['Date'] == yesterday]['New Main Parts'].values[0] if yesterday in daily_counts['Date'].values else 0
                            today_count = daily_counts[daily_counts['Date'] == today_date]['New Main Parts'].values[0] if today_date in daily_counts['Date'].values else 0
                            
                            col1, col2 = st.columns(2)
                            col1.metric("📅 Yesterday's New Main Parts", yesterday_count)
                            col2.metric("📅 Today's New Main Parts", today_count)
                            
                            fig = px.line(daily_counts, x='Date', y='New Main Parts', title=f"Daily New Main Parts for {selected_customer} (Last 60 days)", markers=True)
                            fig.update_layout(xaxis_title="Date", yaxis_title="Number of Main Parts")
                            st.plotly_chart(fig, use_container_width=True)
                            
                            st.subheader("Last 7 Days Breakdown")
                            last_7_days = daily_counts.tail(7)
                            st.dataframe(last_7_days, use_container_width=True)
    
    with tab10:
        st.subheader("🛠️ Programmer Board - Missing Nesting Programs")
        st.caption("Tasks that have no Nesting Number (not programmed yet) for selected departments. Sort by Material (Mtl 10) to prioritize.")
        
        target_depts = ['Laser Cut', 'Laser Tube', 'Punching']   # 修改这里
        available_depts = [d for d in target_depts if d in filtered_df['Current Dept'].unique()]
        if not available_depts:
            st.warning(f"None of the target departments {target_depts} found in data. Available departments: {filtered_df['Current Dept'].unique().tolist()}")
        else:
            selected_depts = st.multiselect("Select departments to include", available_depts, default=available_depts)
            if not selected_depts:
                st.info("Please select at least one department.")
            else:
                mask_dept = filtered_df['Current Dept'].isin(selected_depts)
                mask_no_nesting = filtered_df['Nesting Num'].isna() | (filtered_df['Nesting Num'].astype(str).str.strip() == '')
                programmer_df = filtered_df[mask_dept & mask_no_nesting].copy()
                
                if programmer_df.empty:
                    st.success(f"No missing nesting tasks in the selected departments.")
                else:
                    st.error(f"🚨 {len(programmer_df)} task(s) missing nesting program.")
                    display_cols = ['JobNum/Asm', 'Nesting Num', 'Subpart Part Num', 'Subpart 2D Rev', 'Subpart KK Rev', 'Mtl 10', 'Subpart Qty', 'Current Dept', 'Current Operation', 'Order Category']
                    display_cols = [c for c in display_cols if c in programmer_df.columns]
                    sort_by = st.selectbox("Sort by", options=['Mtl 10', 'JobNum/Asm', 'Subpart Part Num'], index=0)
                    ascending = st.checkbox("Ascending", value=True)
                    programmer_df = programmer_df.sort_values(by=sort_by, ascending=ascending)
                    st.dataframe(programmer_df[display_cols], use_container_width=True, height=500)
                    
                    st.subheader("Summary by Material (Mtl 10)")
                    mtl_counts = programmer_df['Mtl 10'].value_counts().reset_index()
                    mtl_counts.columns = ['Material', 'Count']
                    st.dataframe(mtl_counts, use_container_width=True)
                    
                    with st.expander("Show already programmed tasks (for reference)"):
                        mask_programmed = filtered_df['Current Dept'].isin(selected_depts) & (~filtered_df['Nesting Num'].isna()) & (filtered_df['Nesting Num'].astype(str).str.strip() != '')
                        programmed_df = filtered_df[mask_programmed][display_cols].head(20)
                        st.dataframe(programmed_df, use_container_width=True)
    
    with tab11:
        st.subheader("🛠️ Engineering Workbench Required")
        st.caption("Main parts missing Engineering Workbench: no JobNum/Asm and no Steps in either the main part row or its corresponding subpart row (where Subpart Part Num equals Main Part Num).")
        
        raw_df = st.session_state['df']   # 使用原始数据（未经过滤）
        main_part_nums = raw_df['Main Part Num'].dropna().unique()
        missing_list = []
        for main_part in main_part_nums:
            main_rows = raw_df[raw_df['Main Part Num'] == main_part]
            main_row = main_rows.iloc[0] if not main_rows.empty else None
            sub_rows = raw_df[raw_df['Subpart Part Num'] == main_part]
            sub_row = sub_rows.iloc[0] if not sub_rows.empty else None
            
            has_job = False
            has_steps = False
            
            # 收集负责工程师（优先取子部件行的 Assigned Eng，否则取主部件行）
            assigned_eng = ''
            if sub_row is not None and pd.notna(sub_row.get('Assigned Eng')) and str(sub_row.get('Assigned Eng')).strip() != '':
                assigned_eng = str(sub_row.get('Assigned Eng')).strip()
            elif main_row is not None and pd.notna(main_row.get('Assigned Eng')) and str(main_row.get('Assigned Eng')).strip() != '':
                assigned_eng = str(main_row.get('Assigned Eng')).strip()
            
            if assigned_eng == '':
                assigned_eng = '未分配'
            
            if main_row is not None:
                if pd.notna(main_row.get('JobNum/Asm')) and str(main_row.get('JobNum/Asm')).strip() != '':
                    has_job = True
                for i in range(1, 21):
                    col = f'Step {i}'
                    if col in main_row and pd.notna(main_row[col]) and str(main_row[col]).strip() != '':
                        has_steps = True
                        break
            if sub_row is not None:
                if pd.notna(sub_row.get('JobNum/Asm')) and str(sub_row.get('JobNum/Asm')).strip() != '':
                    has_job = True
                for i in range(1, 21):
                    col = f'Step {i}'
                    if col in sub_row and pd.notna(sub_row[col]) and str(sub_row[col]).strip() != '':
                        has_steps = True
                        break
            if not has_job and not has_steps:
                info_row = main_row if main_row is not None else sub_row
                if info_row is not None:
                    missing_list.append({
                        'Main Part Num': main_part,
                        'Main Part 2D Rev': info_row.get('Main Part 2D Rev', ''),
                        'Main Part 3D Rev': info_row.get('Main Part 3D Rev', ''),
                        'Main Part KK Rev': info_row.get('Main Part KK Rev', ''),
                        'Subpart Part Num': info_row.get('Subpart Part Num', ''),
                        'Assigned Eng': assigned_eng
                    })
        
        missing_eng = pd.DataFrame(missing_list)
        if missing_eng.empty:
            st.success("✅ All main parts have Engineering Workbench completed (JobNum/Asm and Steps exist in either main or subpart row).")
        else:
            st.error(f"🚨 {len(missing_eng)} main part(s) missing Engineering Workbench.")
            display_cols = ['Main Part Num', 'Main Part 2D Rev', 'Main Part 3D Rev', 'Main Part KK Rev', 'Subpart Part Num', 'Assigned Eng']
            display_cols = [c for c in display_cols if c in missing_eng.columns]
            st.dataframe(missing_eng[display_cols], use_container_width=True)
            
            # 按工程师汇总（确保工程师列不为空）
            st.subheader("Summary by Assigned Engineer")
            # 过滤掉空值，只统计有工程师名字的记录
            eng_counts = missing_eng[missing_eng['Assigned Eng'] != '未分配']['Assigned Eng'].value_counts()
            if not eng_counts.empty:
                eng_summary = eng_counts.reset_index()
                eng_summary.columns = ['Engineer', 'Missing Workbench Count']
                st.dataframe(eng_summary, use_container_width=True)
            else:
                st.info("No assigned engineers found for missing workbenches.")
            
            # 按客户汇总（保留原有）
            missing_eng['Customer'] = missing_eng['Main Part Num'].apply(lambda x: x.split('-')[0] if '-' in x else x)
            cust_summary = missing_eng['Customer'].value_counts().reset_index()
            cust_summary.columns = ['Customer', 'Missing Count']
            st.subheader("Summary by Customer")
            st.dataframe(cust_summary, use_container_width=True)

else:
    st.info("👈 Please upload one or more Excel files exported from Epicor (BAQ Report)")
    st.markdown("""
    ### 📌 Instructions
    1. Export BAQ Report from Epicor, ensure the header is on row 6.
    2. Required columns: `Main Part Num`, `Subpart Part Num`, `Step 1`~`Step 20`, `Current Operation`.
    3. Optional: `First Process Plan Date`, `Order Date`, `Exwork Date`, `PO - POLine`, `Order Category`, etc.
    4. You can upload **multiple Excel files** at once (e.g., different customers). The system will combine them.
    5. Use **Complete & Next** buttons in Department Workbench to advance tasks.
    6. **Auto-Calibration**: Enter actual hours (in hours) and click "Calibrate" to adjust future ETAs. Export/Import calibration JSON for persistence.
    7. Download updated Excel to persist progress changes.
    8. Check **Delayed Alerts** tab for overdue tasks.
    9. Use **Job Progress Board** to get an overview of all Jobs and quickly jump to Gantt/Sales views.
    10. Use **Stuck Alerts** to monitor tasks exceeding custom time threshold.
    11. Use **Customer Summary** to analyze main parts by customer and month.
    12. Use **Programmer Board** to see missing nesting tasks.
    13. Use **Engineering WB Required** to see main parts missing engineering workbench.
    14. Use **Order Category Filter** in the sidebar to focus on New Awarded/New Revision items.
    """)
