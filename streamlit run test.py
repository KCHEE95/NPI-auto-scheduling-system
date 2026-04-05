import streamlit as st

steps = ['M-LC-FBR', 'P-DB']
current_idx = 1  # 当前是第二个步骤
remaining_steps = len(steps) - current_idx - 1

html = f"""
<div style="margin-top: 8px;">
    <div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 8px;">
        <span style="font-size: 0.75rem; font-weight: 600; color: #475569;">工序步骤</span>
        <span style="font-size: 0.7rem; color: #64748b;">Step {current_idx+1} of {len(steps)}</span>
    </div>
    <div style="display: flex; gap: 6px; align-items: center;">
        {''.join([f'''
        <div style="
            flex: 1;
            height: 10px;
            background-color: {'#10b981' if i < current_idx else ('#3b82f6' if i == current_idx else '#e2e8f0')};
            border-radius: 20px;
            {'border: 1px solid #3b82f6;' if i == current_idx else ''}
            {'opacity: 0.6;' if i > current_idx else ''}
        " title="{step}"></div>
        ''' for i, step in enumerate(steps)])}
    </div>
    <div style="display: flex; justify-content: space-between; margin-top: 4px; font-size: 0.65rem; color: #94a3b8;">
        <span>开始</span>
        <span>当前</span>
        <span>剩余 {remaining_steps} 步</span>
    </div>
</div>
"""

st.markdown(html, unsafe_allow_html=True)
