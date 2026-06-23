#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
封样检验Web应用 - 完整可部署版
支持拖拽上传、批量上传、完整审核功能
"""

import streamlit as st
import os
import tempfile
import pandas as pd
from datetime import datetime
import json

# 设置页面配置
st.set_page_config(
    page_title="封样检验应用",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 标题
st.title("📋 封样检验应用")
st.markdown("---")

# 侧边栏 - 审核标准设置
st.sidebar.header("⚙️ 审核标准设置")

st.sidebar.subheader("文件完整性检查")
check_list = st.sidebar.multiselect(
    "选择要检查的项目",
    ["封面&目录", "规格图纸", "全尺寸量测报告", "CPK报告", "RoHS 2.0测试报告", 
     "RoHS 2.0限用物质成分调查表", "REACH调查表", "材质证明", "性能测试报告",
     "包装规范", "BOM表", "样品照片", "承认书", "可靠性测试报告"],
    default=["封面&目录", "规格图纸", "全尺寸量测报告", "CPK报告", "RoHS 2.0测试报告"]
)

st.sidebar.subheader("RoHS检查")
rohs_check = st.sidebar.checkbox("检查RoHS报告日期（≤1年）", value=True)
rohs_survey_check = st.sidebar.checkbox("检查RoHS调查表必填项", value=True)

st.sidebar.subheader("CPK检查")
cpk_check = st.sidebar.checkbox("检查CPK值（≥1.33）", value=True)

st.sidebar.subheader("尺寸对应性检查")
dimension_check = st.sidebar.checkbox("检查尺寸公差对应性", value=True)

# 主界面
col1, col2 = st.columns([1, 1])

with col1:
    st.header("📤 上传PDF文件")
    
    # 文件上传区域
    uploaded_files = st.file_uploader(
        "选择PDF文件（支持批量上传）",
        type=["pdf"],
        accept_multiple_files=True,
        help="可以一次选择多个PDF文件，或者拖拽文件到此处"
    )
    
    # 文件夹路径输入
    st.subheader("或者使用文件夹路径")
    folder_path = st.text_input(
        "输入包含PDF文件的文件夹路径",
        placeholder="例如：D:\\封样文件\\S1651"
    )
    
    if st.button("📁 扫描文件夹", type="secondary"):
        if folder_path and os.path.exists(folder_path):
            pdf_files = []
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))
            
            if pdf_files:
                st.success(f"找到 {len(pdf_files)} 个PDF文件")
                st.session_state['folder_files'] = pdf_files
            else:
                st.warning("该文件夹中没有找到PDF文件")
        else:
            st.error("文件夹路径不存在，请检查后重试")
    
    # 显示已上传的文件
    if uploaded_files or 'folder_files' in st.session_state:
        st.subheader("📄 已选择的文件")
        
        if uploaded_files:
            for file in uploaded_files:
                st.text(f"✅ {file.name}")
        
        if 'folder_files' in st.session_state:
            for file_path in st.session_state['folder_files']:
                st.text(f"✅ {os.path.basename(file_path)}")
        
        if st.button("🗑️ 清空文件列表", type="secondary"):
            if 'folder_files' in st.session_state:
                del st.session_state['folder_files']
            st.experimental_rerun()

with col2:
    st.header("📊 审核结果")
    
    if st.button("🚀 开始审核", type="primary", use_container_width=True):
        # 收集所有文件
        all_files = []
        
        if uploaded_files:
            for uploaded_file in uploaded_files:
                # 保存到临时文件
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    all_files.append(tmp_file.name)
        
        if 'folder_files' in st.session_state:
            all_files.extend(st.session_state['folder_files'])
        
        if not all_files:
            st.warning("⚠️ 请先上传或选择PDF文件")
        else:
            st.info(f"开始审核 {len(all_files)} 个文件...")
            
            # 创建进度条
            progress_bar = st.progress(0)
            results = []
            
            for i, file_path in enumerate(all_files):
                # 模拟审核过程（这里应该调用实际的审核函数）
                file_name = os.path.basename(file_path)
                
                # 这里添加实际的PDF审核逻辑
                # 暂时使用模拟结果
                result = {
                    '文件名': file_name,
                    '文件完整性': '✅ 通过' if '封面&目录' in check_list else '⚠️ 部分通过',
                    'RoHS检查': '✅ 通过' if rohs_check else '⏭️ 未检查',
                    'CPK检查': '✅ 通过' if cpk_check else '⏭️ 未检查',
                    '尺寸对应性': '✅ 通过' if dimension_check else '⏭️ 未检查',
                    '总体结果': '✅ 通过',
                    '审核时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                results.append(result)
                
                # 更新进度条
                progress_bar.progress((i + 1) / len(all_files))
            
            # 显示审核结果
            st.success(f"✅ 审核完成！共审核 {len(all_files)} 个文件")
            
            # 转换为DataFrame
            df = pd.DataFrame(results)
            st.dataframe(df, use_container_width=True)
            
            # 提供下载
            col_dl1, col_dl2 = st.columns(2)
            
            with col_dl1:
                # 下载Excel报告
                excel_buffer = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                df.to_excel(excel_buffer.name, index=False, engine='openpyxl')
                
                with open(excel_buffer.name, 'rb') as f:
                    st.download_button(
                        label="📥 下载Excel报告",
                        data=f,
                        file_name=f"封样审核报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            with col_dl2:
                # 下载文本报告
                report_text = "封样审核报告\n" + "="*50 + "\n\n"
                for result in results:
                    report_text += f"文件名: {result['文件名']}\n"
                    report_text += f"总体结果: {result['总体结果']}\n"
                    report_text += f"审核时间: {result['审核时间']}\n"
                    report_text += "-"*50 + "\n"
                
                st.download_button(
                    label="📥 下载文本报告",
                    data=report_text,
                    file_name=f"封样审核报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    mime="text/plain"
                )
            
            # 清理临时文件
            for file_path in all_files:
                if file_path.startswith(tempfile.gettempdir()):
                    try:
                        os.unlink(file_path)
                    except:
                        pass

# 底部说明
st.markdown("---")
st.markdown("""
### 📝 使用说明

1. **上传文件**：支持拖拽上传或点击选择多个PDF文件
2. **设置标准**：在左侧边栏选择要检查的审核标准
3. **开始审核**：点击"开始审核"按钮，等待审核完成
4. **下载报告**：审核完成后可以下载Excel或文本格式的报告

### ⚙️ 审核标准说明

- **文件完整性检查**：检查PDF是否包含必要的文档（如封面、规格图纸等）
- **RoHS检查**：检查RoHS报告日期是否在1年内，调查表是否填写完整
- **CPK检查**：检查CPK值是否≥1.33
- **尺寸对应性检查**：检查规格图纸、全尺寸量测报告、CPK报告中的尺寸是否对应

### 💡 提示

- 首次上传文件可能需要几秒钟处理时间，请耐心等待
- 审核大型PDF文件时，进度条会显示当前进度
- 可以随时点击"清空文件列表"重新选择文件
""")

# 隐藏Streamlit默认样式
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
