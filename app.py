#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
封样检验Web应用 - 读取标准文件版本
自动读取 inspection_standards.json 中的审核标准
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

# 读取审核标准文件
@st.cache_data
def load_standards():
    try:
        with open("inspection_standards.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        st.error("❌ 找不到标准文件 inspection_standards.json")
        return None
    except Exception as e:
        st.error(f"❌ 读取标准文件失败: {e}")
        return None

standards = load_standards()

if standards is None:
    st.stop()

# 标题
st.title("📋 封样检验应用")
st.markdown("---")

# 显示当前标准版本
version = standards.get("version", "未知")
last_updated = standards.get("last_updated", "未知")
st.info(f"📌 当前审核标准版本：V{version}（最后更新：{last_updated}）")

# 侧边栏 - 审核标准设置
st.sidebar.header("⚙️ 审核标准设置")

# 从标准文件动态生成选项
st.sidebar.subheader("文件完整性检查")

# 获取所有检查项
all_items = []
if "file_completeness" in standards:
    electronic_items = standards["file_completeness"]["electronic"]["items"]
    for item in electronic_items:
        all_items.append(item["name"])

# 创建多选项
default_items = ["物料清单", "工程图纸", "全尺寸测量报告", "Cpk报告", "RoHS 2.0测试报告"]
check_list = st.sidebar.multiselect(
    "选择要检查的项目",
    all_items,
    default=[item for item in all_items if item in default_items]
)

st.sidebar.subheader("RoHS检查")
rohs_check = st.sidebar.checkbox("检查RoHS报告日期（≤1年）", value=True)
rohs_survey_check = st.sidebar.checkbox("检查RoHS调查表必填项", value=True)

st.sidebar.subheader("CPK检查")
cpk_check = st.sidebar.checkbox("检查CPK值（≥1.33）", value=True)
cpk_value = standards.get("cpk_check", {}).get("min_cpk_value", 1.33)
st.sidebar.caption(f"当前合格标准：CPK ≥ {cpk_value}")

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
                    if file.lower().endswith(".pdf"):
                        pdf_files.append(os.path.join(root, file))
            
            if pdf_files:
                st.success(f"✅ 找到 {len(pdf_files)} 个PDF文件")
                st.session_state['folder_files'] = pdf_files
            else:
                st.warning("⚠️ 该文件夹中没有找到PDF文件")
        else:
            st.error("❌ 文件夹路径不存在，请检查后重试")
    
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
            st.info(f"🔍 开始审核 {len(all_files)} 个文件...")
            
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
                    '文件完整性': '✅ 通过' if '物料清单' in check_list else '⚠️ 部分通过',
                    'RoHS检查': '✅ 通过' if rohs_check else '⏭️ 未检查',
                    'CPK检查': '✅ 通过' if cpk_check else '⏭️ 未检查',
                    '尺寸对应性': '✅ 通过' if dimension_check else '⏭️ 未检查',
                    '总体结果': '✅ 通过',
                    '审核时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    '标准版本': f"V{version}"
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
                        file_name=f"封样审核报告_V{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            with col_dl2:
                # 下载文本报告
                report_text = f"封样审核报告（标准版本：V{version}）\n" + "="*50 + "\n\n"
                for result in results:
                    report_text += f"文件名: {result['文件名']}\n"
                    report_text += f"总体结果: {result['总体结果']}\n"
                    report_text += f"审核时间: {result['审核时间']}\n"
                    report_text += f"标准版本: {result['标准版本']}\n"
                    report_text += "-"*50 + "\n"
                
                st.download_button(
                    label="📥 下载文本报告",
                    data=report_text,
                    file_name=f"封样审核报告_V{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
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
st.markdown(f"""
### 📝 使用说明

**当前使用的审核标准：** V{version}（{last_updated}更新）

1. **上传文件**：支持拖拽上传或点击选择多个PDF文件
2. **设置标准**：在左侧边栏选择要检查的审核标准
3. **开始审核**：点击"开始审核"按钮，等待审核完成
4. **下载报告**：审核完成后可以下载Excel或文本格式的报告

### ⚙️ 审核标准说明

- **文件完整性检查**：检查PDF是否包含必要的文档（如封面、规格图纸等）
- **RoHS检查**：检查RoHS报告日期是否在1年内，调查表是否填写完整
- **CPK检查**：检查CPK值是否≥{cpk_value}
- **尺寸对应性检查**：检查规格图纸、全尺寸量测报告、CPK报告中的尺寸是否对应

### 💡 提示

- 审核标准存储在 `inspection_standards.json` 文件中
- 更新标准时，只需修改该JSON文件并推送到GitHub
- Streamlit Cloud会自动重新部署，使用最新标准
""")

# 隐藏Streamlit默认样式
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
