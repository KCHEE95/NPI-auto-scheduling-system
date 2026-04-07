# System Prompt: AI Auto Scheduling & Progress Tracking System

## 1. System Name
AI Auto Scheduling & Progress Tracking System

## 2. System Objectives
- Automatically parse NPI work order data from Epicor ERP exported Excel files.
- Provide real-time progress visualization, estimated finish dates (ETA), capacity load analysis, delay/stuck alerts, and engineering work status dashboards for production, sales, programmers, and engineers.
- Support manual operation advancement ("Complete & Next") and automatic calibration of operation lead times based on actual hours feedback.
- Enable multi-user collaboration (3-6 users) via Streamlit Cloud or on-premises deployment.

## 3. Data Source & Input Format
- Source: Epicor BAQ Report Excel file – header is on row 6 (index 5).
- Mandatory columns:
  - Main Part Num
  - Subpart Part Num
  - Step 1 ... Step 20 (or Step1...Step20)
  - Current Operation
  - JobNum/Asm (format e.g., 525651-0; -0 = main part, -1, -2 = subparts)
  - Nesting Num (empty = not programmed)
  - Exwork Date (delivery date)
  - Order Date
  - Order Category (New Awarded, New Revision, Repeated Order)
  - PO - POLine
  - Mtl 10 (material info)
  - Assigned Eng
- Optional columns: First Process Plan Date, Subpart Qty, Subpart 2D Rev, Subpart KK Rev, etc.

## 4. Core Business Logic

### 4.1 Operation Chain Parsing
- Extract non-empty values from Step 1 to Step 20 -> ordered list _steps.
- Map each operation code to a department using OP_TO_DEPT dictionary (e.g., P-DB -> Deburr).

### 4.2 ETA Calculation
- Based on current operation position in _steps, sum standard lead times (LEAD_TIME dictionary, unit: days, based on 10 working hours per day) of remaining operations.
- If current operation is empty or not in _steps, sum all steps.
- Outsourced operation F-NPV1 fixed at 7 days.
- Main part (-0) ETA = max(ETA of all its subparts) + remaining days of main part itself.

### 4.3 Special Rule for Main Part ETA
- Main part must wait for all subparts to finish before its own remaining operations.
- Hence main part ETA = max(subpart ETA) + main part remaining days.

### 4.4 Department Capacity Load
- Each department has a maximum concurrent task number (DEPT_CAPACITY, integer).
- Load % = (current task count / capacity) * 100%.

### 4.5 Manual Operation Advancement ("Complete & Next")
- In Department Workbench, each task card has a "Complete & Next" button.
- Click -> current operation advances to next step in _steps, recalc ETA, log change (change_log).
- If last step -> mark as COMPLETED.

### 4.6 Auto-Calibration
- Operator inputs actual hours in the task card and clicks "Calibrate".
- Exponential smoothing: new_standard_days = 0.7 * old_standard_days + 0.3 * (actual_hours / 10).
- Calibrated values stored in st.session_state.lead_time_override; can be exported/imported as JSON.

### 4.7 Delayed Alerts
- Status = "⚠️ Delayed" if ETA < today, else "✅ On track".
- Dedicated tab shows all delayed tasks, department breakdown, and days delayed.

### 4.8 Stuck Alerts
- Monitors only tasks advanced via "Complete & Next" (have _step_start_time).
- User-definable threshold (hours). Exceeding threshold -> "🔴 Stuck".
- Shows stayed days, threshold, exceed ratio.

### 4.9 Programmer Board
- Shows tasks in departments ['Laser Cut', 'Laser Tube', 'Punching'] with empty Nesting Num.
- Supports sorting by material (Mtl 10) and material summary.

### 4.10 Engineering WB Required Board
- Lists main parts that have neither JobNum/Asm nor any Step column content (checked in both the main part row and its corresponding subpart row where Subpart Part Num equals the main part number).
- Helps engineers identify missing engineering work (drawing, routing creation).

### 4.11 Customer Summary
- Aggregates main parts only (-0) by month based on Exwork Date.
- Daily trend chart based on Order Date (last 60 days) shows new main parts per day.

### 4.12 Sales Query
- Search by Job number, PO number, or subpart part number.
- Shows summary (total subparts, on-track/delayed count, bottleneck department, Exwork Date) and filterable subpart detail table.

### 4.13 Gantt Chart (per Job)
- For a selected Job, displays timeline from Planned Date to ETA for all subparts.
- Bar color = current department; red dashed line = today.

## 5. Technology Stack & Deployment
- Framework: Streamlit (Python)
- Data processing: Pandas, NumPy
- Visualisation: Plotly
- State management: st.session_state
- Deployment: Streamlit Cloud (public) or on-premises server (internal network)
- Multi-file upload: supports simultaneous upload of multiple Excel files, automatically merged.

## 6. User Interface & Interaction
- Sidebar:
  - Password login (default admin123)
  - Auto-calibration controls (export/import JSON, reset)
  - Order Category multi-select filter (default New Awarded & New Revision)
  - Change log export/clear
- 11 Tabs:
  1. All Items
  2. Department Workbench (card-style tasks with progress bar, Complete & Next, Calibration)
  3. Capacity Dashboard
  4. Sales Query
  5. Job Gantt Chart
  6. Delayed Alerts
  7. Job Progress Board (global Job overview with quick jumps)
  8. Stuck Alerts (customisable threshold)
  9. Customer Summary (main parts only)
  10. Programmer Board (missing nesting tasks)
  11. Engineering WB Required (missing engineering work)

## 7. Data Persistence & Export
- Users can download an updated Excel containing manually advanced progress and calibrated values.
- Calibration data can be exported as JSON and reloaded.
- Change log can be exported as JSON backup.

## 8. Known Limitations & Notes
- No write-back to Epicor; progress updates require manual export and re-import (or API integration).
- All data resides in memory after upload; page refresh loses data (re-upload needed).
- Multiple users have independent session_state – changes are not shared.
- Cleaning logic (clean_fake_started_jobs) clears Current Operation of main parts that have subparts but no real progress, to avoid false "first step" display.

## 9. Extensibility Suggestions
- Integrate Epicor REST API for bidirectional sync.
- Add role-based access control (different customers per user).
- Implement PuLP linear programming for finite capacity planning.
