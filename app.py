import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px

# ========== Password protection ==========
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        pwd = st.sidebar.text_input("Enter system password", type="password")
        if pwd == "admin123":   # ← change this password as needed
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
st.caption("Auto-parsed from Epicor BAQ Report | Supports operation chain & ETA calculation")

# ========== Configuration (modify according to your actual data) ==========
LEAD_TIME = {
    'W-CDS-A': 1.0, 'W-LWD': 1.0, 'M-LC-FBR': 2.0, 'P-DB': 0.5,
    'M-BD': 1.5, 'P-GRD': 1.0, 'P-DGR': 0.8, 'P-MK-A': 0.5,
    'F-PT': 0.3, 'P-DMK-A': 0.4, 'F-INK': 0.2, '2-PK-A': 0.3,
    'N-MC': 0.7, 'P-TU-A': 0.6, 'D-TAP-A': 0.4, 'P-PCKLNG': 0.5,
    'F-NPV1': 0.8, 'ASSY-A': 1.0, 'P-BF': 0.4, 'C-SAW': 0.6,
    'DEFAULT': 1.0
}

OP_TO_DEPT = {
    'W-CDS-A': 'Cutting', 'W-LWD': 'Laser Welding', 'M-LC-FBR': 'Machining',
    'P-DB': 'Drilling', 'M-BD': 'Bending', 'P-GRD': 'Grinding',
    'P-DGR': 'Deburring', 'P-MK-A': 'Marking', 'F-PT': 'Painting',
    'P-DMK-A': 'Dispensing', 'F-INK': 'Printing', '2-PK-A': 'Packing A',
    'N-MC': 'CNC', 'P-TU-A': 'Tapping', 'D-TAP-A': 'Threading',
    'P-PCKLNG': 'Packing', 'F-NPV1': 'Post-treatment 1', 'ASSY-A': 'Assembly',
    'P-BF': 'Stamping', 'C-SAW': 'Sawing', 'DEFAULT': 'Unassigned'
}

DEPT_CAPACITY = {
    'Cutting': 5, 'Laser Welding': 3, 'Machining': 8, 'Drilling': 4,
    'Bending': 3, 'Grinding': 2, 'Deburring': 2, 'Marking': 2,
    'Painting': 1, 'Dispensing': 2, 'Printing': 1, 'Packing A': 3,
    'CNC': 4, 'Tapping': 2, 'Threading': 2, 'Packing': 4,
    'Post-treatment 1': 2, 'Assembly': 3, 'Stamping': 3, 'Sawing': 2,
}

# ========== Helper functions ==========
@st.cache_data
def load_excel(file):
    """Read Excel, skip first 5 rows, header at row 6, filter valid subparts"""
    df = pd.read_excel(file, header=5)
    df = df.dropna(how='all')
    # Forward fill Main Part Num
    if 'Main Part Num' in df.columns:
        df['Main Part Num'] = df['Main Part Num'].ffill()
    else:
        st.error("Excel missing 'Main Part Num' column")
        st.stop()
    
    # Keep only rows with non-empty Subpart Part Num
    if 'Subpart Part Num' in df.columns:
        df = df[df['Subpart Part Num'].notna() & (df['Subpart Part Num'] != '')]
    else:
        st.error("Excel missing 'Subpart Part Num' column")
        st.stop()
    
    # Ensure JobNum/Asm and Nesting Num columns exist (create empty if missing)
    if 'JobNum/Asm' not in df.columns:
        df['JobNum/Asm'] = ''
    if 'Nesting Num' not in df.columns:
        df['Nesting Num'] = ''
    
    return df

def extract_step_sequence(row):
    """Extract operation list from Step columns (supports 'Step 1' or 'Step1')"""
    steps = []
    step_col_candidates = [f'Step {i}' for i in range(1, 21)]
    if step_col_candidates[0] not in row.index:
        step_col_candidates = [f'Step{i}' for i in range(1, 21)]
    
    for col in step_col_candidates:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != '':
            steps.append(row[col])
    return steps

def compute_eta(row, today):
    """Calculate ETA based on current operation and step chain"""
    current_op = row.get('Current Operation')
    steps = row['_steps']
    
    if not steps:
        return today + timedelta(days=7)
    
    if pd.isna(current_op) or current_op == '' or current_op not in steps:
        remaining_days = sum(LEAD_TIME.get(op, LEAD_TIME['DEFAULT']) for op in steps)
    else:
        try:
            idx = steps.index(current_op)
        except ValueError:
            idx = -1
        remaining_days = 0
        for op in steps[idx+1:]:
            remaining_days += LEAD_TIME.get(op, LEAD_TIME['DEFAULT'])
    
    remaining_days = max(remaining_days, 0.5)
    return today + timedelta(days=remaining_days)

def get_dept_from_op(op):
    if pd.isna(op) or op == '':
        return 'Unassigned'
    return OP_TO_DEPT.get(op, OP_TO_DEPT['DEFAULT'])

# ========== Main interface ==========
uploaded_file = st.sidebar.file_uploader("📁 Upload Excel file exported from Epicor", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = load_excel(uploaded_file)
    st.sidebar.success(f"✅ Loaded {len(df)} valid subparts")
    
    # Extract step sequences
    df['_steps'] = df.apply(extract_step_sequence, axis=1)
    
    # Calculate ETA
    today = datetime.now().date()
    df['ETA'] = df.apply(lambda row: compute_eta(row, today), axis=1)
    
    # Map to department
    df['Current Dept'] = df['Current Operation'].apply(get_dept_from_op)
    
    # Status flag
    df['Status'] = df['ETA'].apply(lambda x: '⚠️ Delayed' if x < today else '✅ On track')
    
    # ========== Multi-page tabs ==========
    tab1, tab2, tab3, tab4 = st.tabs(["📋 All Items", "🏭 Department Workbench", "📈 Capacity Dashboard", "🔍 Sales Query"])
    
    with tab1:
        st.subheader("Real-time status of all subparts")
        display_cols = ['Main Part Num', 'Subpart Part Num', 'JobNum/Asm', 'Nesting Num', 
                        'Current Operation', 'Current Dept', 'ETA', 'Status', 'Assigned Eng']
        display_cols = [c for c in display_cols if c in df.columns]
        df_display = df[display_cols].sort_values('ETA')
        st.dataframe(df_display, use_container_width=True, height=500)
        
        with st.expander("🔍 View full operation chain for each subpart"):
            for _, row in df.iterrows():
                if row['_steps']:
                    steps_str = " → ".join(row['_steps'])
                    st.write(f"**{row['Subpart Part Num']}** (Job: {row['JobNum/Asm']}, Nest: {row['Nesting Num']}) : {steps_str}")
    
    with tab2:
        st.subheader("Department to-do list")
        dept_list = sorted(df['Current Dept'].unique())
        selected_dept = st.selectbox("Select department", dept_list)
        dept_cols = ['Main Part Num', 'Subpart Part Num', 'JobNum/Asm', 'Nesting Num', 
                     'Current Operation', 'ETA', 'Status', 'Assigned Eng']
        dept_cols = [c for c in dept_cols if c in df.columns]
        dept_df = df[df['Current Dept'] == selected_dept][dept_cols].sort_values('ETA')
        st.dataframe(dept_df, use_container_width=True)
        
        overdue = dept_df[dept_df['Status'] == '⚠️ Delayed']
        if not overdue.empty:
            st.warning(f"⚠️ {len(overdue)} potentially delayed task(s) in this department")
    
    with tab3:
        st.subheader("Department capacity load")
        dept_load = df['Current Dept'].value_counts().reset_index()
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
        search_term = st.text_input("Enter Main Part Num, Subpart Part Num, or JobNum/Asm (supports partial match)")
        if search_term:
            mask = (df['Main Part Num'].str.contains(search_term, case=False, na=False) |
                    df['Subpart Part Num'].str.contains(search_term, case=False, na=False) |
                    df['JobNum/Asm'].astype(str).str.contains(search_term, case=False, na=False))
            result = df[mask]
            if not result.empty:
                for _, row in result.iterrows():
                    eta_str = row['ETA'].strftime('%Y-%m-%d') if pd.notna(row['ETA']) else 'Unknown'
                    st.info(f"**{row['Subpart Part Num']}**  \n"
                            f"- JobNum/Asm: {row['JobNum/Asm']}  \n"
                            f"- Nesting Num: {row['Nesting Num']}  \n"
                            f"- Current Operation: {row['Current Operation']}  \n"
                            f"- Department: {row['Current Dept']}  \n"
                            f"- Estimated Completion Date: {eta_str}  \n"
                            f"- Status: {row['Status']}")
            else:
                st.warning("No matching Part or Job found")
else:
    st.info("👈 Please upload the Excel file exported from Epicor (BAQ Report)")
    st.markdown("""
    ### 📌 Instructions
    1. Export BAQ Report from Epicor, ensure the header is on row 6 (code handles this automatically)
    2. Must include columns: `Main Part Num`, `Subpart Part Num`, `Step 1`~`Step 20` (or `Step1`~`Step20`), `Current Operation`
    3. Strongly recommended to include `JobNum/Asm` and `Nesting Num` for production use
    4. After upload, the system will:
       - Filter out empty subpart rows
       - Extract full operation chain for each subpart
       - Calculate ETA based on current operation and standard lead times
       - Show tasks by department and capacity load
    """)
