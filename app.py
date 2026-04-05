import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
from io import BytesIO

# ========== Password protection ==========
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
st.caption("Auto-parsed from Epicor BAQ Report | Supports operation chain, ETA, and task completion")

# ========== Configuration ==========
LEAD_TIME = {
    'W-CDS-A': 1.0, 'W-LWD': 1.0, 'M-LC-FBR': 2.0, 'P-DB': 0.5,
    'M-BD': 1.5, 'P-GRD': 1.0, 'P-DGR': 0.8, 'P-MK-A': 0.5,
    'F-PT': 0.3, 'P-DMK-A': 0.4, 'F-INK': 0.2, '2-PK-A': 0.3,
    'N-MC': 0.7, 'P-TU-A': 0.6, 'D-TAP-A': 0.4, 'P-PCKLNG': 0.5,
    'F-NPV1': 0.8, 'ASSY-A': 1.0, 'P-BF': 0.4, 'C-SAW': 0.6,
    'DEFAULT': 1.0
}

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
    df = pd.read_excel(file, header=5)
    df = df.dropna(how='all')
    if 'Main Part Num' in df.columns:
        df['Main Part Num'] = df['Main Part Num'].ffill()
    else:
        st.error("Excel missing 'Main Part Num' column")
        st.stop()
    
    if 'Subpart Part Num' in df.columns:
        df = df[df['Subpart Part Num'].notna() & (df['Subpart Part Num'] != '')]
    else:
        st.error("Excel missing 'Subpart Part Num' column")
        st.stop()
    
    # Ensure required columns exist
    for col in ['JobNum/Asm', 'Nesting Num', 'Exwork Date', 'Subpart Qty',
                'Subpart 2D Rev', 'Subpart KK Rev', 'Mtl 10', 'Subpart Part Image']:
        if col not in df.columns:
            df[col] = ''
    return df

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
    """返回当前工序的下一个工序，如果已完成或没有下一步则返回空字符串"""
    if not steps:
        return ''
    if pd.isna(current_op) or current_op == '':
        return steps[0] if steps else ''  # 没有当前工序，则第一步是下一步
    if current_op not in steps:
        return ''
    idx = steps.index(current_op)
    if idx + 1 < len(steps):
        return steps[idx + 1]
    else:
        return 'COMPLETED'  # 已是最后一步

def compute_eta(row, today):
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

def update_task_to_next_operation(df, index):
    """将指定行的任务推进到下一个工序，同时更新 Next Operation 和 ETA"""
    row = df.loc[index]
    steps = row['_steps']
    current_op = row['Current Operation']
    if pd.isna(current_op) or current_op == '':
        # 没有当前工序，无法推进
        return df, False, "No current operation"
    if current_op not in steps:
        return df, False, f"Current operation '{current_op}' not found in step chain"
    current_idx = steps.index(current_op)
    if current_idx + 1 >= len(steps):
        # 已经是最后一步，标记为完成
        df.at[index, 'Current Operation'] = 'COMPLETED'
        df.at[index, 'Current Dept'] = 'Completed'
        df.at[index, 'Next Operation'] = ''
        # 重新计算 ETA（已完成，ETA 设为今天或空）
        df.at[index, 'ETA'] = datetime.now().date()
        return df, True, "Task completed (no next operation)"
    else:
        next_op = steps[current_idx + 1]
        df.at[index, 'Current Operation'] = next_op
        df.at[index, 'Current Dept'] = get_dept_from_op(next_op)
        # 更新 Next Operation
        df.at[index, 'Next Operation'] = get_next_operation(next_op, steps)
        # 重新计算 ETA
        df.at[index, 'ETA'] = compute_eta(df.loc[index], datetime.now().date())
        return df, True, f"Moved to next operation: {next_op}"

# ========== Main interface ==========
uploaded_file = st.sidebar.file_uploader("📁 Upload Excel file exported from Epicor", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Load data (cached, but we will use session_state for modifications)
    if 'original_df' not in st.session_state or st.sidebar.button("Reload original file"):
        df = load_excel(uploaded_file)
        # Extract step sequences
        df['_steps'] = df.apply(extract_step_sequence, axis=1)
        # Calculate initial ETA
        today = datetime.now().date()
        df['ETA'] = df.apply(lambda row: compute_eta(row, today), axis=1)
        df['Current Dept'] = df['Current Operation'].apply(get_dept_from_op)
        df['Status'] = df['ETA'].apply(lambda x: '⚠️ Delayed' if x < today else '✅ On track')
        # Add Next Operation column
        df['Next Operation'] = df.apply(lambda row: get_next_operation(row['Current Operation'], row['_steps']), axis=1)
        # Convert Exwork Date
        if 'Exwork Date' in df.columns:
            df['Exwork Date'] = pd.to_datetime(df['Exwork Date'], errors='coerce')
        st.session_state['original_df'] = df
        st.session_state['df'] = df.copy()
        st.session_state['file_name'] = uploaded_file.name
        st.rerun()
    else:
        df = st.session_state['df']
    
    st.sidebar.success(f"✅ Loaded {len(df)} valid subparts")
    
    # ========== Tabs ==========
    tab1, tab2, tab3, tab4 = st.tabs(["📋 All Items", "🏭 Department Workbench", "📈 Capacity Dashboard", "🔍 Sales Query"])
    
    with tab1:
        st.subheader("Real-time status of all subparts")
        base_cols = ['Main Part Num', 'Subpart Part Num', 'JobNum/Asm', 'Nesting Num',
                     'Current Operation', 'Next Operation', 'Current Dept', 'ETA', 'Status', 'Assigned Eng']
        extra_cols = ['Exwork Date', 'Subpart Qty', 'Subpart 2D Rev', 'Subpart KK Rev', 'Mtl 10']
        display_cols = [c for c in base_cols + extra_cols if c in df.columns]
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
        selected_dept = st.selectbox("Select department", dept_list, key="dept_select")
        
        col1, col2 = st.columns(2)
        with col1:
            job_filter = st.text_input("Filter by JobNum/Asm (partial match)", key="job_filter")
        with col2:
            nest_filter = st.text_input("Filter by Nesting Num (partial match)", key="nest_filter")
        
        filtered_df = df[df['Current Dept'] == selected_dept]
        if job_filter:
            filtered_df = filtered_df[filtered_df['JobNum/Asm'].astype(str).str.contains(job_filter, case=False, na=False)]
        if nest_filter:
            filtered_df = filtered_df[filtered_df['Nesting Num'].astype(str).str.contains(nest_filter, case=False, na=False)]
        
        if filtered_df.empty:
            st.info("No tasks match the filters.")
        else:
            cols_to_show = ['Main Part Num', 'Subpart Part Num', 'JobNum/Asm', 'Nesting Num',
                            'Current Operation', 'Next Operation', 'ETA', 'Status', 'Assigned Eng',
                            'Exwork Date', 'Subpart Qty', 'Mtl 10']
            cols_to_show = [c for c in cols_to_show if c in filtered_df.columns]
            
            for idx, row in filtered_df.iterrows():
                with st.container():
                    col_a, col_b = st.columns([0.85, 0.15])
                    with col_a:
                        row_data = {col: row[col] for col in cols_to_show}
                        st.write(pd.DataFrame([row_data]).T)
                    with col_b:
                        if st.button(f"✅ Complete & Next", key=f"complete_{idx}"):
                            updated_df, success, message = update_task_to_next_operation(st.session_state['df'], idx)
                            if success:
                                st.session_state['df'] = updated_df
                                st.success(f"Task {row['Subpart Part Num']}: {message}")
                                st.rerun()
                            else:
                                st.error(f"Failed: {message}")
                st.divider()
        
        if st.button("📥 Download updated Excel (with progress changes)"):
            output_df = st.session_state['df'].drop(columns=['_steps'], errors='ignore')
            # Convert datetime columns to string for Excel compatibility
            for col in ['ETA', 'Exwork Date']:
                if col in output_df.columns:
                    output_df[col] = output_df[col].astype(str)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='UpdatedSchedule')
            st.download_button(
                label="Download Excel",
                data=output.getvalue(),
                file_name=f"updated_{st.session_state['file_name']}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
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
                    exwork_str = row['Exwork Date'].strftime('%Y-%m-%d') if pd.notna(row.get('Exwork Date')) else 'Not set'
                    st.info(f"**{row['Subpart Part Num']}**  \n"
                            f"- JobNum/Asm: {row['JobNum/Asm']}  \n"
                            f"- Nesting Num: {row['Nesting Num']}  \n"
                            f"- Current Operation: {row['Current Operation']}  \n"
                            f"- Next Operation: {row.get('Next Operation', '')}  \n"
                            f"- Department: {row['Current Dept']}  \n"
                            f"- Estimated Completion Date: {eta_str}  \n"
                            f"- Exwork Date (Delivery): {exwork_str}  \n"
                            f"- Subpart Qty: {row.get('Subpart Qty', '')}  \n"
                            f"- Material: {row.get('Mtl 10', '')}  \n"
                            f"- Status: {row['Status']}")
            else:
                st.warning("No matching Part or Job found")
else:
    st.info("👈 Please upload the Excel file exported from Epicor (BAQ Report)")
    st.markdown("""
    ### 📌 Instructions
    1. Export BAQ Report from Epicor, ensure the header is on row 6 (code handles this automatically)
    2. Must include columns: `Main Part Num`, `Subpart Part Num`, `Step 1`~`Step 20` (or `Step1`~`Step20`), `Current Operation`
    3. The system will automatically compute `Next Operation` based on the step chain.
    4. In **Department Workbench**, use filters and click **Complete & Next** to advance tasks. The `Next Operation` will update correctly.
    5. Download the updated Excel to persist changes.
    """)
