#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit 入口包装器 - 全局异常捕获
用于在 Streamlit Cloud 上显示详细错误信息
"""

import sys
import os
import traceback

# 设置工作目录为脚本所在目录（Streamlit Cloud 需要）
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 先导入 streamlit 并配置页面
import streamlit as st

try:
    st.set_page_config(
        page_title="封样检验应用",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded"
    )
except Exception:
    pass  # set_page_config 可能重复调用

# 导入并运行真正的应用逻辑
try:
    # 将当前目录加入 Python 路径
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

    # 导入主应用模块（执行其中的全部代码）
    import importlib.util
    spec = importlib.util.spec_from_file_location("seal_app", os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_real.py"))
    seal_app = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(seal_app)

except Exception as e:
    st.error(f"❌ 应用启动失败：{type(e).__name__}: {e}")
    st.code(traceback.format_exc())
    st.markdown("---")
    st.markdown("### 🔧 环境调试信息")
    st.text(f"Python版本: {sys.version}")
    st.text(f"工作目录: {os.getcwd()}")
    try:
        files = os.listdir('.')
        st.text(f"目录文件 ({len(files)}个): {', '.join(files[:20])}")
    except Exception as e2:
        st.text(f"无法列出文件: {e2}")
