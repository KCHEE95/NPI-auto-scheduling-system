# 系统描述 Prompt：AI 自动排产与进度跟踪系统

## 1. 系统名称
AI 自动排产与进度跟踪系统（AI Auto Scheduling & Progress Tracking System）

## 2. 系统目标
- 从 Epicor ERP 导出的 Excel 文件中自动解析 NPI 项目的工单数据。
- 为生产部门、销售、程序员、工程师提供实时进度可视化、预计完成日期（ETA）、产能负载分析、延期/卡住警报、以及工程工作状态看板。
- 支持手动推进工序（Complete & Next）和工序耗时自动校准（基于实际工时反馈）。
- 支持多用户（3-6人）协同，通过 Streamlit Cloud 或本地部署访问。

## 3. 数据源与输入格式
- 来源：Epicor ERP 导出的 BAQ Report Excel 文件，表头位于第6行（索引5）。
- 关键列（必须包含）：
  - Main Part Num（主部件编号）
  - Subpart Part Num（子部件编号）
  - Step 1 至 Step 20（工序步骤链，也可能为 Step1 格式）
  - Current Operation（当前工序代码）
  - JobNum/Asm（工单编号，格式如 525651-0，-0 为主部件，-1、-2 等为子部件）
  - Nesting Num（套料程序编号，为空表示未编程）
  - Exwork Date（出货日期）
  - Order Date（订单日期）
  - Order Category（订单类别：New Awarded, New Revision, Repeated Order）
  - PO - POLine（采购订单号）
  - Mtl 10（材料信息）
  - Assigned Eng（负责工程师）
- 可选列：First Process Plan Date, Subpart Qty, Subpart 2D Rev, Subpart KK Rev 等。

## 4. 核心业务逻辑

### 4.1 工序链解析
- 从 Step 1 到 Step 20 提取非空值，形成有序列表 _steps。
- 每个工序代码映射到实际部门（OP_TO_DEPT 字典），例如 P-DB -> Deburr。

### 4.2 ETA 计算
- 根据当前工序在 _steps 中的位置，累加剩余所有工序的标准耗时（LEAD_TIME 字典，单位：天，基于每天10工作小时换算）。
- 如果当前工序为空或不在 _steps 中，则累加全部步骤。
- 外包工序 F-NPV1 固定 7 天。
- 主部件（-0）的 ETA = max(所有子部件的 ETA) + 主部件自身剩余天数。

### 4.3 主部件 ETA 特殊规则
- 主部件必须等待所有子部件完成后才开始自身的剩余工序。
- 因此主部件 ETA = max(子部件 ETA) + 主部件剩余天数。

### 4.4 部门产能负载
- 每个部门有最大并发任务数 DEPT_CAPACITY（整数）。
- 负载率 = (当前任务数 / 容量) * 100%。

### 4.5 手动推进工序（Complete & Next）
- 在部门工作台中，每个任务卡片有“Complete & Next”按钮。
- 点击后，当前工序变为步骤链中的下一步，重新计算 ETA，并记录变更日志（change_log）。
- 如果已是最后一步，则标记为 COMPLETED。

### 4.6 自动校准（Calibration）
- 操作员可在任务卡片中输入实际加工小时数，点击“Calibrate”。
- 系统使用指数平滑：新标准天数 = 0.7 * 旧标准天数 + 0.3 * (实际小时数 / 10)。
- 校准值保存在 st.session_state.lead_time_override 中，可导出/导入 JSON 文件。

### 4.7 延期警报（Delayed Alerts）
- 状态判断：ETA < 今天 则为“⚠️ Delayed”，否则“✅ On track”。
- 单独标签页展示所有延期任务，按部门统计，并显示已延期天数。

### 4.8 卡住警报（Stuck Alerts）
- 只监控通过 Complete & Next 推进过的任务（有 _step_start_time）。
- 用户可自定义阈值（小时数），超过阈值则标记为“🔴 Stuck”。
- 展示停留天数、标准阈值、超出比例。

### 4.9 程序员看板（Programmer Board）
- 只显示当前部门为 ['Laser Cut', 'Laser Tube', 'Punching'] 且 Nesting Num 为空的任务。
- 支持按材料（Mtl 10）排序，提供材料汇总。

### 4.10 工程看板（Engineering WB Required）
- 显示所有主部件中，既没有 JobNum/Asm 也没有任何 Step 列内容的部件（包括主部件行本身和对应的子部件行）。
- 用于提醒工程师未完成工程工作（画图、添加工序）。

### 4.11 客户汇总（Customer Summary）
- 基于 Exwork Date 按月份聚合主部件（-0）数量，展示每个客户的月度分布和总计。
- 每日趋势图基于 Order Date 显示最近60天新增主部件数量。

### 4.12 销售查询（Sales Query）
- 支持按 Job 编号、PO 号、子部件编号搜索。
- 展示汇总信息（总子部件数、按时/延期数、瓶颈部门、出货日期等）和子部件详情表格（可筛选）。

### 4.13 甘特图（Job Gantt Chart）
- 为选定的 Job 展示所有子部件的计划开始（Planned Date）到预计完成（ETA）的时间线。
- 条形颜色按当前部门区分，红色虚线标记今天。

## 5. 技术栈与部署
- 前端/后端：Streamlit（Python）
- 数据处理：Pandas, NumPy
- 可视化：Plotly
- 状态管理：st.session_state
- 部署方式：Streamlit Cloud（公网）或本地服务器（内网）
- 多文件上传：支持同时上传多个 Excel 文件，自动合并

## 6. 用户交互与界面
- 侧边栏：
  - 密码登录（默认 admin123）
  - 自动校准配置（导出/导入 JSON，重置）
  - Order Category 多选筛选（默认只选 New Awarded 和 New Revision）
  - 变更日志导出/清除
- 11个标签页：
  1. All Items（所有子部件状态表格）
  2. Department Workbench（部门工作台，卡片式任务列表，含进度条、完成按钮、校准）
  3. Capacity Dashboard（产能负载柱状图及表格）
  4. Sales Query（销售查询）
  5. Job Gantt Chart（甘特图）
  6. Delayed Alerts（延期警报）
  7. Job Progress Board（全局Job进度看板，含跳转）
  8. Stuck Alerts（卡住警报，自定义阈值）
  9. Customer Summary（客户汇总，仅主部件）
  10. Programmer Board（程序员看板，缺少 Nesting 的任务）
  11. Engineering WB Required（工程看板，缺少工程工作的主部件）

## 7. 数据持久化与导出
- 用户可通过“Download updated Excel”导出包含手动推进进度和校准值的 Excel。
- 校准数据可导出为 JSON，下次上传时加载。
- 变更日志可导出为 JSON 备份。

## 8. 已知限制与注意事项
- 系统不自动写回 Epicor，进度更新需手动导出 Excel 再导入 Epicor（或通过 API 集成）。
- 单次上传文件合并后，所有数据驻留在内存中，刷新页面需重新上传。
- 多用户同时操作时，各自的 session_state 独立，变更不共享。
- 清洗逻辑（clean_fake_started_jobs）会将“有子部件但无任何进度的主部件”的 Current Operation 清空，避免虚假第一步。

## 9. 扩展性建议
- 可集成 Epicor REST API 实现双向同步。
- 可增加用户角色权限（不同用户只看自己负责的客户）。
- 可增加 PuLP 线性规划实现有限能力排产。
