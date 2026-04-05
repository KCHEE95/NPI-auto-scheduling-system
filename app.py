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
    'F-NPV1': 7.0,      # 外包工序，固定7天
    'DEFAULT': 1.0      # 未知工序默认1天（10小时）
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
    """计算单个任务基于当前工序和工序链的预计完成日期（不考虑等待子部件）"""
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

def compute_remaining_days(row, today):
    """返回当前任务从当前工序到结束所需的总天数（浮点数）"""
    current_op = row.get('Current Operation')
    steps = row['_steps']
    if not steps:
        return 7.0
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
    """
    对于每个 Job，找到主部件（-0），计算：
    主部件最终 ETA = max(所有子部件 ETA) + 主部件自身剩余加工天数
    如果主部件自身已完成（Current Operation = COMPLETED），则不加天数。
    """
    df['_job_base'] = df['JobNum/Asm'].apply(get_job_base)
    df['_is_main'] = df['JobNum/Asm'].astype(str).str.endswith('-0')
    
    # 计算每个 Job 基础编号下，所有子部件（非主部件）的 ETA 最大值
    subpart_max_eta = {}
    for job_base in df['_job_base'].unique():
        sub_df = df[(df['_job_base'] == job_base) & (~df['_is_main'])]
        if not sub_df.empty:
            subpart_max_eta[job_base] = sub_df['ETA'].max()
    
    # 更新主部件行
    for idx, row in df[df['_is_main']].iterrows():
        job_base = row['_job_base']
        # 主部件自身剩余天数
        remaining_days = compute_remaining_days(row, today)
        # 子部件最晚日期
        sub_max = subpart_max_eta.get(job_base, pd.NaT)
        if pd.notna(sub_max) and row['Current Operation'] != 'COMPLETED':
            # 主部件必须等待子部件完成后才开始自身剩余工序
            new_eta = sub_max + timedelta(days=remaining_days)
        else:
            # 没有子部件或主部件已完成，则使用自身 ETA
            new_eta = row['ETA']  # 自身 ETA 已基于今日+剩余天数计算
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
    df.at[index, 'Status'] = '✅ On track' if df.at[index, 'ETA'] >= today else '⚠️ Delayed'
    
    # 重新计算该 Job 的主部件 ETA（基于更新后的子部件）
    job_base = get_job_base(df.at[index, 'JobNum/Asm'])
    if job_base:
        job_mask = df['_job_base'] == job_base
        main_mask = job_mask & df['_is_main']
        if main_mask.any():
            main_idx = df[main_mask].index[0]
            # 重新计算子部件最大 ETA（排除主部件自身）
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

# ========== Main interface ==========
uploaded_file = st.sidebar.file_uploader("📁 Upload Excel file exported from Epicor", type=["xlsx", "xls"])

if uploaded_file is not None:
    if 'original_df' not in st.session_state or st.sidebar.button("Reload original file"):
        df = load_excel(uploaded_file)
        df['_steps'] = df.apply(extract_step_sequence, axis=1)
        today = datetime.now().date()
        # 先计算每个任务自身的 ETA（基于当前工序）
        df['ETA'] = df.apply(lambda row: compute_eta(row, today), axis=1)
        df['Current Dept'] = df['Current Operation'].apply(get_dept_from_op)
        df['Next Operation'] = df.apply(lambda row: get_next_operation(row['Current Operation'], row['_steps']), axis=1)
        df['Planned Date'] = df.apply(get_planned_date, axis=1)
        df['_job_base'] = df['JobNum/Asm'].apply(get_job_base)
        df['_is_main'] = df['JobNum/Asm'].astype(str).str.endswith('-0')
        df['Status'] = df['ETA'].apply(lambda x: '✅ On track' if x >= today else '⚠️ Delayed')
        # 更新主部件的 ETA（基于子部件最晚日期 + 自身剩余天数）
        df = update_main_part_eta(df, today)
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
        st.caption("**Status explanation**: ✅ On track = Estimated finish date is today or in the future; ⚠️ Delayed = Estimated finish date has passed but task not completed.\n\n**Note**: For main parts (JobNum/Asm ending with -0), the Est. Finish Date = latest finish date among all subparts + remaining processing time of the main part itself.")
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
        st.info("💡 **JobNum/Asm format**: `-0` indicates the main part; `-1`, `-2` etc. indicate subparts. Main part's Est. Finish Date = max(subpart ETA) + main part's own remaining days.")
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
    3. The system automatically computes `Next Operation`, `Planned Date`, and `Est. Finish Date` (ETA).
    4. **For main parts (JobNum/Asm ending with -0)**, the Est. Finish Date = latest finish date among all subparts + remaining processing time of the main part itself.
    5. **Status**: ✅ On track = Est. Finish Date is today or in the future; ⚠️ Delayed = Est. Finish Date has passed.
    6. In **Department Workbench**, use filters and click **Complete & Next** to advance tasks. The main part's ETA will automatically update.
    7. Download the updated Excel to persist changes.
    """)
