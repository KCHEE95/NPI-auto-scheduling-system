import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
from io import BytesIO
import re

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
                'Subpart 2D Rev', 'Subpart KK Rev', 'Mtl 10', 'Subpart Part Image',
                'First Process Plan Date', 'Order Date']:
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

def get_planned_date(row):
    if 'First Process Plan Date' in row and pd.notna(row['First Process Plan Date']) and row['First Process Plan Date'] != '':
        return pd.to_datetime(row['First Process Plan Date'], errors='coerce')
    elif 'Order Date' in row and pd.notna(row['Order Date']) and row['Order Date'] != '':
        return pd.to_datetime(row['Order Date'], errors='coerce')
    else:
        return pd.NaT

def get_job_base(job_num):
    """从 JobNum/Asm 提取基础编号，例如 '525651-0' -> '525651'"""
    if pd.isna(job_num) or job_num == '':
        return ''
    # 使用正则去掉最后的 -数字
    match = re.match(r'^([^-]+)', str(job_num))
    return match.group(1) if match else str(job_num)

def update_main_part_eta(df, today):
    """
    对于每个 Job 基础编号，找到主部件（JobNum/Asm 以 '-0' 结尾的行），
    将其 ETA 更新为该 Job 下所有行的最大 ETA。
    同时更新其 Status 和 Current Dept（如果主部件当前工序不是 COMPLETED，部门保持不变）。
    """
    # 添加辅助列：Job 基础编号
    df['_job_base'] = df['JobNum/Asm'].apply(get_job_base)
    # 标记是否为主部件（以 -0 结尾）
    df['_is_main'] = df['JobNum/Asm'].astype(str).str.endswith('-0')
    
    # 计算每个 Job 基础编号的最大 ETA
    group_max_eta = df.groupby('_job_base')['ETA'].max()
    
    # 更新主部件行的 ETA 和 Status
    for idx, row in df[df['_is_main']].iterrows():
        job_base = row['_job_base']
        if job_base in group_max_eta:
            new_eta = group_max_eta[job_base]
            df.at[idx, 'ETA'] = new_eta
            df.at[idx, 'Status'] = '✅ On track' if new_eta >= today else '⚠️ Delayed'
            # 注意：主部件的部门可能要根据其 Current Operation 重新映射（不变）
    return df

def update_task_to_next_operation(df, index, today):
    """推进任务，并重新计算相关主部件的 ETA"""
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
    
    # 更新该任务的状态
    df.at[index, 'Status'] = '✅ On track' if df.at[index, 'ETA'] >= today else '⚠️ Delayed'
    
    # 重新计算该任务所属 Job 的主部件 ETA（如果有）
    job_base = get_job_base(df.at[index, 'JobNum/Asm'])
    if job_base:
        # 找到同一 Job 下的所有行
        job_mask = df['_job_base'] == job_base
        # 重新计算该组的最大 ETA
        max_eta = df.loc[job_mask, 'ETA'].max()
        # 找到主部件行（_is_main = True）并更新
        main_mask = job_mask & df['_is_main']
        if main_mask.any():
            main_idx = df[main_mask].index[0]
            df.at[main_idx, 'ETA'] = max_eta
            df.at[main_idx, 'Status'] = '✅ On track' if max_eta >= today else '⚠️ Delayed'
    
    return df, True, f"Moved to next operation: {next_op if current_idx+1 < len(steps) else 'COMPLETED'}"

# ========== Main interface ==========
uploaded_file = st.sidebar.file_uploader("📁 Upload Excel file exported from Epicor", type=["xlsx", "xls"])

if uploaded_file is not None:
    if 'original_df' not in st.session_state or st.sidebar.button("Reload original file"):
        df = load_excel(uploaded_file)
        df['_steps'] = df.apply(extract_step_sequence, axis=1)
        today = datetime.now().date()
        df['ETA'] = df.apply(lambda row: compute_eta(row, today), axis=1)
        df['Current Dept'] = df['Current Operation'].apply(get_dept_from_op)
        df['Next Operation'] = df.apply(lambda row: get_next_operation(row['Current Operation'], row['_steps']), axis=1)
        df['Planned Date'] = df.apply(get_planned_date, axis=1)
        # 添加辅助列
        df['_job_base'] = df['JobNum/Asm'].apply(get_job_base)
        df['_is_main'] = df['JobNum/Asm'].astype(str).str.endswith('-0')
        # 先计算初始状态（基于自身工序）
        df['Status'] = df['ETA'].apply(lambda x: '✅ On track' if x >= today else '⚠️ Delayed')
        # 然后更新主部件的 ETA（基于同 Job 所有行的最大 ETA）
        df = update_main_part_eta(df, today)
        # 转换 Exwork Date
        if 'Exwork Date' in df.columns:
            df['Exwork Date'] = pd.to_datetime(df['Exwork Date'], errors='coerce')
        st.session_state['original_df'] = df
        st.session_state['df'] = df.copy()
        st.session_state['file_name'] = uploaded_file.name
        st.rerun()
    else:
        df = st.session_state['df']
    
    st.sidebar.success(f"✅ Loaded {len(df)} valid subparts")
    
    tab1, tab2, tab3, tab4 = st.tabs(["📋 All Items", "🏭 Department Workbench", "📈 Capacity Dashboard", "🔍 Sales Query"])
    
    with tab1:
        st.subheader("Real-time status of all subparts")
        st.caption("**Status explanation**: ✅ On track = Estimated finish date is today or in the future; ⚠️ Delayed = Estimated finish date has passed but task not completed.\n\n**Note**: For main parts (JobNum/Asm ending with -0), the Est. Finish Date is calculated as the latest finish date among all its subparts.")
        base_cols = ['Main Part Num', 'Subpart Part Num', 'JobNum/Asm', 'Nesting Num',
                     'Current Operation', 'Next Operation', 'Current Dept', 
                     'Planned Date', 'ETA', 'Status', 'Assigned Eng']
        extra_cols = ['Exwork Date', 'Subpart Qty', 'Subpart 2D Rev', 'Subpart KK Rev', 'Mtl 10']
        display_cols = [c for c in base_cols + extra_cols if c in df.columns]
        df_display = df[display_cols].rename(columns={'ETA': 'Est. Finish Date'}).sort_values('Est. Finish Date')
        st.dataframe(df_display, use_container_width=True, height=500)
        
        with st.expander("🔍 View full operation chain for each subpart"):
            for _, row in df.iterrows():
                if row['_steps']:
                    steps_str = " → ".join(row['_steps'])
                    st.write(f"**{row['Subpart Part Num']}** (Job: {row['JobNum/Asm']}, Nest: {row['Nesting Num']}) : {steps_str}")
    
    with tab2:
        st.subheader("Department to-do list")
        st.info("💡 **JobNum/Asm format**: `-0` indicates the main part; `-1`, `-2` etc. indicate subparts of the same job number. Main part's Est. Finish Date is auto-calculated as the latest among its subparts.")
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
                            'Current Operation', 'Next Operation', 'Planned Date', 'ETA', 'Status', 'Assigned Eng',
                            'Exwork Date', 'Subpart Qty', 'Mtl 10']
            cols_to_show = [c for c in cols_to_show if c in filtered_df.columns]
            for idx, row in filtered_df.iterrows():
                with st.container():
                    col_a, col_b = st.columns([0.85, 0.15])
                    with col_a:
                        row_data = {col: row[col] for col in cols_to_show}
                        if 'ETA' in row_data:
                            row_data['Est. Finish Date'] = row_data.pop('ETA').strftime('%Y-%m-%d') if pd.notna(row_data['ETA']) else ''
                        if 'Planned Date' in row_data and pd.notna(row_data['Planned Date']):
                            row_data['Planned Date'] = row_data['Planned Date'].strftime('%Y-%m-%d')
                        st.write(pd.DataFrame([row_data]).T)
                    with col_b:
                        if st.button(f"✅ Complete & Next", key=f"complete_{idx}"):
                            today = datetime.now().date()
                            updated_df, success, message = update_task_to_next_operation(st.session_state['df'], idx, today)
                            if success:
                                st.session_state['df'] = updated_df
                                st.success(f"Task {row['Subpart Part Num']}: {message}")
                                st.rerun()
                            else:
                                st.error(f"Failed: {message}")
                st.divider()
        
        if st.button("📥 Download updated Excel (with progress changes)"):
            output_df = st.session_state['df'].drop(columns=['_steps', '_job_base', '_is_main'], errors='ignore')
            for col in ['ETA', 'Exwork Date', 'Planned Date']:
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
                    planned_str = row['Planned Date'].strftime('%Y-%m-%d') if pd.notna(row.get('Planned Date')) else 'Not set'
                    exwork_str = row['Exwork Date'].strftime('%Y-%m-%d') if pd.notna(row.get('Exwork Date')) else 'Not set'
                    st.info(f"**{row['Subpart Part Num']}**  \n"
                            f"- JobNum/Asm: {row['JobNum/Asm']}  \n"
                            f"- Nesting Num: {row['Nesting Num']}  \n"
                            f"- Planned Date: {planned_str}  \n"
                            f"- Current Operation: {row['Current Operation']}  \n"
                            f"- Next Operation: {row.get('Next Operation', '')}  \n"
                            f"- Department: {row['Current Dept']}  \n"
                            f"- Est. Finish Date: {eta_str}  \n"
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
    3. The system automatically computes `Next Operation`, `Planned Date` (from `First Process Plan Date` or `Order Date`), and `Est. Finish Date` (ETA).
    4. **For main parts (JobNum/Asm ending with -0)**, the Est. Finish Date is calculated as the latest finish date among all its subparts (including itself).
    5. **Status**: ✅ On track = Est. Finish Date is today or in the future; ⚠️ Delayed = Est. Finish Date has passed.
    6. In **Department Workbench**, use filters and click **Complete & Next** to advance tasks. The main part's ETA will automatically update.
    7. Download the updated Excel to persist changes.
    """)
