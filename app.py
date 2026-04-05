import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px

# ========== 简单密码保护 ==========
def check_password():
    """返回 True 表示验证通过"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        pwd = st.sidebar.text_input("请输入系统密码", type="password")
        if pwd == "admin123":   # ← 你可以修改这个密码
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.sidebar.error("密码错误")
            return False
    return True

# 只有验证通过后才显示主界面
if not check_password():
    st.stop()

# ========== 页面配置 ==========
st.set_page_config(page_title="AI 自动排产系统", layout="wide")
st.title("📊 AI 自动排产与进度跟踪系统")
st.caption("基于 Epicor BAQ Report 自动解析 | 支持工序链 & ETA 计算")

# ========== 1. 配置区（请根据实际修改）==========
# 工序代码 -> 标准耗时（天）—— 你需要根据实际补充
LEAD_TIME = {
    'W-CDS-A': 1.0, 'W-LWD': 1.0, 'M-LC-FBR': 2.0, 'P-DB': 0.5,
    'M-BD': 1.5, 'P-GRD': 1.0, 'P-DGR': 0.8, 'P-MK-A': 0.5,
    'F-PT': 0.3, 'P-DMK-A': 0.4, 'F-INK': 0.2, '2-PK-A': 0.3,
    'N-MC': 0.7, 'P-TU-A': 0.6, 'D-TAP-A': 0.4, 'P-PCKLNG': 0.5,
    'F-NPV1': 0.8, 'ASSY-A': 1.0, 'P-BF': 0.4, 'C-SAW': 0.6,
    'DEFAULT': 1.0
}

# 工序代码 -> 实际部门名称
OP_TO_DEPT = {
    'W-CDS-A': '切割部', 'W-LWD': '激光焊接部', 'M-LC-FBR': '加工部',
    'P-DB': '钻孔部', 'M-BD': '弯板部', 'P-GRD': '研磨部',
    'P-DGR': '去毛刺部', 'P-MK-A': '标记部', 'F-PT': '喷涂部',
    'P-DMK-A': '点胶部', 'F-INK': '印刷部', '2-PK-A': '包装A组',
    'N-MC': '数控部', 'P-TU-A': '攻牙部', 'D-TAP-A': '攻丝部',
    'P-PCKLNG': '包装部', 'F-NPV1': '后处理1组', 'ASSY-A': '装配部',
    'P-BF': '冲压部', 'C-SAW': '锯切部', 'DEFAULT': '待分配'
}

# 部门最大并发任务数（用于负载率计算，可选）
DEPT_CAPACITY = {
    '切割部': 5, '激光焊接部': 3, '加工部': 8, '钻孔部': 4,
    '弯板部': 3, '研磨部': 2, '去毛刺部': 2, '标记部': 2,
    '喷涂部': 1, '点胶部': 2, '印刷部': 1, '包装A组': 3,
    '数控部': 4, '攻牙部': 2, '攻丝部': 2, '包装部': 4,
    '后处理1组': 2, '装配部': 3, '冲压部': 3, '锯切部': 2,
}

# ========== 2. 辅助函数 ==========
@st.cache_data
def load_excel(file):
    df = pd.read_excel(file, header=5)   # header在第6行
    df = df.dropna(how='all')
    df['Main Part Num'] = df['Main Part Num'].fillna(method='ffill')
    return df

def extract_step_sequence(row):
    steps = []
    for i in range(1, 21):
        col = f'Step {i}'
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != '':
            steps.append(row[col])
    return steps

def compute_eta(row, today):
    current_op = row['Current Operation']
    steps = row['_steps']
    if not steps:
        return today + timedelta(days=7)
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
    return OP_TO_DEPT.get(op, OP_TO_DEPT['DEFAULT'])

# ========== 3. 主界面 ==========
uploaded_file = st.sidebar.file_uploader("📁 上传 Epicor 导出的 Excel 文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = load_excel(uploaded_file)
    st.sidebar.success(f"✅ 加载 {len(df)} 条记录")

    df['_steps'] = df.apply(extract_step_sequence, axis=1)
    today = datetime.now().date()
    df['ETA'] = df.apply(lambda row: compute_eta(row, today), axis=1)
    df['Current Dept'] = df['Current Operation'].apply(get_dept_from_op)
    df['Status'] = df['ETA'].apply(lambda x: '⚠️ 可能延期' if x < today else '✅ 正常')

    tab1, tab2, tab3, tab4 = st.tabs(["📋 所有项目", "🏭 部门工作台", "📈 产能仪表板", "🔍 销售查询"])

    with tab1:
        st.subheader("所有子部件实时状态")
        cols = ['Main Part Num', 'Subpart Part Num', 'Current Operation', 'Current Dept', 'ETA', 'Status', 'Assigned Eng']
        cols = [c for c in cols if c in df.columns]
        st.dataframe(df[cols], use_container_width=True, height=500)

    with tab2:
        dept_list = sorted(df['Current Dept'].unique())
        sel = st.selectbox("选择部门", dept_list)
        dept_df = df[df['Current Dept'] == sel][['Main Part Num', 'Subpart Part Num', 'Current Operation', 'ETA', 'Status']]
        st.dataframe(dept_df, use_container_width=True)

    with tab3:
        dept_load = df['Current Dept'].value_counts().reset_index()
        dept_load.columns = ['部门', '任务数']
        dept_load['容量'] = dept_load['部门'].map(DEPT_CAPACITY).fillna(5)
        dept_load['负载率 (%)'] = (dept_load['任务数'] / dept_load['容量'] * 100).round(1)
        fig = px.bar(dept_load, x='部门', y='任务数', color='负载率 (%)', title='各部门负载')
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(dept_load)

    with tab4:
        q = st.text_input("输入 Main Part Num 或 Subpart Part Num")
        if q:
            mask = df['Main Part Num'].str.contains(q, case=False, na=False) | df['Subpart Part Num'].str.contains(q, case=False, na=False)
            res = df[mask]
            if not res.empty:
                for _, r in res.iterrows():
                    st.info(f"**{r['Subpart Part Num']}** → 当前工序: {r['Current Operation']}  |  ETA: {r['ETA'].strftime('%Y-%m-%d')}  |  {r['Status']}")
            else:
                st.warning("未找到")
else:
    st.info("👈 请从左侧上传 Excel 文件（BAQ Report）")
