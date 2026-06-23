#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""封样检验应用 - 诊断版"""
import streamlit as st
import os
import sys

try:
    st.set_page_config(
        page_title="封样检验应用",
        page_icon="📋",
        layout="wide",
    )
except:
    pass

st.title("📋 封样检验应用 - 诊断模式")
st.markdown("---")

# 显示环境信息
st.subheader("🔧 环境信息")
st.text(f"Python: {sys.version}")
st.text(f"工作目录: {os.getcwd()}")
st.text(f"文件列表: {os.listdir('.')}")

# 测试导入
st.subheader("📦 依赖测试")
imports_ok = []
imports_fail = []
for mod_name in ["pandas", "pdfplumber", "openpyxl", "json", "re", "datetime"]:
    try:
        __import__(mod_name)
        imports_ok.append(mod_name)
    except Exception as e:
        imports_fail.append(f"{mod_name}: {e}")

if imports_ok:
    st.success(f"✅ 正常导入: {', '.join(imports_ok)}")
if imports_fail:
    st.error(f"❌ 导入失败: {'; '.join(imports_fail)}")

# 测试标准文件读取
st.subheader("📄 标准文件测试")
try:
    import json
    with open("inspection_standards.json", "r", encoding="utf-8") as f:
        standards = json.load(f)
    st.success(f"✅ 标准文件加载成功 (版本: {standards.get('version', '未知')})")
except Exception as e:
    st.error(f"❌ 标准文件加载失败: {e}")

st.markdown("---")
st.info("💡 如果您看到此页面，说明 Streamlit Cloud 部署成功！点击下方按钮切换到完整版。")

if st.button("🔄 切换到完整版", type="primary"):
    st.info("请等待管理员将完整版代码推送...")
