import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px

# ========== 密码保护 ==========
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        pwd = st.sidebar.text_input("请输入系统密码", type="password")
        if pwd == "admin123":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.sidebar.error("密码错误")
            return False
    return True

if not check_password():
    st.stop()

st.set_page_config(page_title="AI 自动排产系统", layout="wide")
st.title("📊 AI 自动排产与进度跟踪系统")
st.caption("基于 Epicor BAQ Report 自动解析 | 支持工序链 & ETA 计算")

# ========== 配置区（请根据实际修改）==========
LEAD_TIME = {
    'W-CDS-A': 1.0, 'W-LWD': 1.0, 'M-LC-FBR': 2.0, 'P-DB': 0.5,
    'M-BD': 1.5, 'P-GRD': 1.0, 'P-DGR': 0.8, 'P-MK-A': 0.5,
    'F-PT': 0.3, 'P-DMK-A': 0.4, 'F-INK': 0.2, '2-PK-A': 0.3,
    'N-MC': 0.7, 'P-TU-A': 0.6, 'D-TAP-A': 0.4, 'P-PCKLNG': 0.5,
    'F-NPV1': 0.8, 'ASSY-A': 1.0, 'P-BF': 0.4, 'C-SAW': 0.6,
    'DEFAULT': 1.0
}

OP_TO_DEPT = {
    'W-CDS-A': '切割部', 'W-LWD': '激光焊接部', 'M-LC-FBR': '加工部',
    'P-DB': '钻孔部', 'M-BD': '弯板部', 'P-GRD': '研磨部',
    'P-DGR': '去毛刺部', 'P-MK-A': '标记部', 'F-PT': '喷涂部',
    'P-DMK-A': '点胶部', 'F-INK': '印刷部', '2-PK-A': '包装A组',
    'N-MC': '数控部', 'P-TU-A': '攻牙部', 'D-TAP-A': '攻丝部',
    'P-PCKLNG': '包装部', 'F-NPV1': '后处理1组', 'ASSY-A': '装配部',
    'P-BF': '冲压部', 'C-SAW': '锯切部', 'DEFAULT': '待分配'
}

DEPT_CAPACITY = {
    '切割部': 5, '激光焊接部': 3, '加工部': 8, '钻孔部': 4,
    '弯板部': 3, '研磨部': 2, '去毛刺部': 2, '标记部': 2,
    '喷涂部': 1, '点胶部': 2, '印刷部': 1, '包装A组': 3,
    '数控部': 4, '攻牙部': 2, '攻丝部': 2, '包装部': 4,
    '后处理1组': 2, '装配部': 3, '冲压部': 3, '锯切部': 2,
}

# ========== 辅助函数 ==========
@st.cache_data
def load_excel(file):
    """读取Excel，跳过前5行，第6行为列名，并过滤有效子部件行"""
    df = pd.read_excel(file, header=5)
    df = df.dropna(how='all')
    # 向下填充 Main Part Num
    if 'Main Part Num' in df.columns:
        df['Main Part Num'] = df['Main Part Num'].ffill()
    else:
        st.error("Excel 缺少 'Main Part Num' 列")
        st.stop()
    
    # 关键：只保留 Subpart Part Num 非空的行（真正的子部件）
    if 'Subpart Part Num' in df.columns:
        df = df[df['Subpart Part Num'].notna() & (df['Subpart Part Num'] != '')]
    else:
        st.error("Excel 缺少 'Subpart Part Num' 列")
        st.stop()
    
    return df

def extract_step_sequence(row):
    """从 Step 列提取工序列表，兼容 'Step 1' 和 'Step1' 两种列名"""
    steps = []
    # 先判断列名是 'Step 1' 还是 'Step1'
    step_col_candidates = [f'Step {i}' for i in range(1, 21)]
    # 检查第一个候选列是否存在
    if step_col_candidates[0] not in row.index:
        # 尝试无空格版本
        step_col_candidates = [f'Step{i}' for i in range(1, 21)]
    
    for col in step_col_candidates:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != '':
            steps.append(row[col])
    return steps

def compute_eta(row, today):
    """根据当前工序和步骤链计算 ETA"""
    current_op = row.get('Current Operation')
    steps = row['_steps']
    
    # 没有步骤链：给默认 7 天
    if not steps:
        return today + timedelta(days=7)
    
    # 当前工序为空或不在步骤链中：从第一步开始算全部步骤
    if pd.isna(current_op) or current_op == '' or current_op not in steps:
        remaining_days = sum(LEAD_TIME.get(op, LEAD_TIME['DEFAULT']) for op in steps)
    else:
        # 找到当前工序的位置，累加后续步骤
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
        return '待分配'
    return OP_TO_DEPT.get(op, OP_TO_DEPT['DEFAULT'])

# ========== 主界面 ==========
uploaded_file = st.sidebar.file_uploader("📁 上传 Epicor 导出的 Excel 文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = load_excel(uploaded_file)
    st.sidebar.success(f"✅ 加载 {len(df)} 个有效子部件")
    
    # 提取工序链
    df['_steps'] = df.apply(extract_step_sequence, axis=1)
    
    # 计算 ETA
    today = datetime.now().date()
    df['ETA'] = df.apply(lambda row: compute_eta(row, today), axis=1)
    
    # 映射部门
    df['Current Dept'] = df['Current Operation'].apply(get_dept_from_op)
    
    # 状态标记（基于 ETA 与今天比较）
    df['Status'] = df['ETA'].apply(lambda x: '⚠️ 可能延期' if x < today else '✅ 正常')
    
    # ========== 多页面 ==========
    tab1, tab2, tab3, tab4 = st.tabs(["📋 所有项目", "🏭 部门工作台", "📈 产能仪表板", "🔍 销售查询"])
    
    with tab1:
        st.subheader("所有子部件实时状态")
        display_cols = ['Main Part Num', 'Subpart Part Num', 'Current Operation', 'Current Dept', 'ETA', 'Status', 'Assigned Eng']
        display_cols = [c for c in display_cols if c in df.columns]
        # 按 ETA 排序，紧急的在前
        df_display = df[display_cols].sort_values('ETA')
        st.dataframe(df_display, use_container_width=True, height=500)
        
        # 显示工序链示例（可展开）
        with st.expander("🔍 查看每个子部件的完整工序链"):
            for _, row in df.iterrows():
                if row['_steps']:
                    steps_str = " → ".join(row['_steps'])
                    st.write(f"**{row['Subpart Part Num']}** : {steps_str}")
    
    with tab2:
        st.subheader("按部门查看待办事项")
        dept_list = sorted(df['Current Dept'].unique())
        selected_dept = st.selectbox("选择部门", dept_list)
        dept_df = df[df['Current Dept'] == selected_dept][
            ['Main Part Num', 'Subpart Part Num', 'Current Operation', 'ETA', 'Status', 'Assigned Eng']
        ].sort_values('ETA')
        st.dataframe(dept_df, use_container_width=True)
        
        # 超期提醒
        overdue = dept_df[dept_df['Status'] == '⚠️ 可能延期']
        if not overdue.empty:
            st.warning(f"⚠️ 该部门有 {len(overdue)} 个可能延期的任务")
    
    with tab3:
        st.subheader("部门产能负载")
        dept_load = df['Current Dept'].value_counts().reset_index()
        dept_load.columns = ['部门', '任务数']
        dept_load['容量'] = dept_load['部门'].map(DEPT_CAPACITY).fillna(5)
        dept_load['负载率 (%)'] = (dept_load['任务数'] / dept_load['容量'] * 100).round(1)
        dept_load = dept_load.sort_values('负载率 (%)', ascending=False)
        
        fig = px.bar(dept_load, x='部门', y='任务数', color='负载率 (%)',
                     title='各部门当前任务负载（颜色越深越忙）',
                     labels={'任务数': '当前任务数'})
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(dept_load, use_container_width=True)
        
        overload = dept_load[dept_load['负载率 (%)'] > 100]
        if not overload.empty:
            st.error(f"⚠️ 以下部门已超负荷：{', '.join(overload['部门'].tolist())}")
    
    with tab4:
        st.subheader("销售快速查询")
        search_term = st.text_input("输入 Main Part Num 或 Subpart Part Num（支持模糊匹配）")
        if search_term:
            mask = df['Main Part Num'].str.contains(search_term, case=False, na=False) | \
                   df['Subpart Part Num'].str.contains(search_term, case=False, na=False)
            result = df[mask]
            if not result.empty:
                for _, row in result.iterrows():
                    eta_str = row['ETA'].strftime('%Y-%m-%d') if pd.notna(row['ETA']) else '未知'
                    st.info(f"**{row['Subpart Part Num']}**  \n"
                            f"- 当前工序: {row['Current Operation']}  \n"
                            f"- 所在部门: {row['Current Dept']}  \n"
                            f"- 预计完成日期: {eta_str}  \n"
                            f"- 状态: {row['Status']}")
            else:
                st.warning("未找到匹配的 Part")
else:
    st.info("👈 请从左侧上传 Epicor 导出的 Excel 文件（BAQ Report）")
    st.markdown("""
    ### 📌 使用说明
    1. 从 Epicor 导出 BAQ Report，确保表头在第 6 行（代码已自动处理）
    2. 必须包含列：`Main Part Num`, `Subpart Part Num`, `Step 1` ~ `Step 20`（或 `Step1`~`Step20`）, `Current Operation`
    3. 上传后系统会自动：
       - 过滤掉没有子部件编号的空行
       - 提取每个子部件的完整工序链
       - 根据当前工序和标准工时计算 ETA
       - 按部门展示任务和产能负载
    """)
