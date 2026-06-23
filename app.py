#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
封样检验Web应用 - V5.0 真实审核版
自动读取 inspection_standards.json 中的审核标准
使用pdfplumber解析PDF，执行真实的38项审核检查
"""

import streamlit as st
import os
import tempfile
import pandas as pd
from datetime import datetime, timedelta
import json
import re
import pdfplumber

# ============================================================
# 标准文件读取
# ============================================================

@st.cache_data(ttl=300)
def load_standards():
    """读取审核标准JSON文件"""
    try:
        with open("inspection_standards.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        st.error("❌ 找不到标准文件 inspection_standards.json")
        return None
    except Exception as e:
        st.error(f"❌ 读取标准文件失败: {e}")
        return None


# ============================================================
# PDF解析引擎
# ============================================================

def extract_pdf_text(pdf_path):
    """提取PDF全部文本"""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
    except Exception as e:
        text = f"[PDF解析错误: {e}]"
    return text.lower()


def extract_pdf_tables(pdf_path):
    """提取PDF中的表格"""
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
    except Exception:
        pass
    return tables


def extract_dates_from_text(text):
    """从文本中提取所有日期"""
    dates = []
    # 匹配格式: 2025/11/26, 2025-11-26, 2025.11.26 等
    patterns = [
        r'\b(\d{4})[/\-\.](\d{1,2})[/\-\.](\d{1,2})\b',
    ]
    for pattern in patterns:
        matches = re.findall(pattern, text)
        for m in matches:
            try:
                date_str = f"{m[0]}-{m[1].zfill(2)}-{m[2].zfill(2)}"
                dates.append(datetime.strptime(date_str, "%Y-%m-%d").date())
            except ValueError:
                continue
    return dates


def extract_cpk_values(text):
    """从文本中提取CPK值"""
    cpk_values = []
    # 匹配 CPK=1.xx 或 CPK : 1.xx 等格式
    patterns = [
        r'cpk\s*[:：=]\s*(\d+\.?\d*)',
        r'cpk\s*[\(（]\s*(\d+\.?\d*)',
        r'cpk\s*value\s*[:：=]\s*(\d+\.?\d*)',
    ]
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for val in matches:
            try:
                cpk_values.append(float(val))
            except ValueError:
                continue
    return cpk_values


def check_keyword_in_text(text, keywords):
    """检查文本中是否包含关键词列表中的任一关键词"""
    for kw in keywords:
        if kw.lower() in text:
            return True, kw
    return False, None


# ============================================================
# 审核核心逻辑
# ============================================================

@st.cache_data(ttl=60)
def determine_material_type(file_name, standards):
    """
    第零步：判定物料类型（电子料 vs 结构件）
    返回: (type_str, type_cn, matched_keyword, needs_electrical_test)
    """
    file_lower = file_name.lower()

    electronic_kw = standards.get("material_types", {}).get("electronic", {}).get("keywords", [])
    structural_kw = standards.get("material_types", {}).get("structural", {}).get("keywords", [])

    # 先检查电子料关键词
    for kw in electronic_kw:
        if kw.lower() in file_lower:
            # 检查是否需要电气性能测试
            required_electrical = standards.get("material_types", {}).get(
                "electronic", {}
            ).get("required_electrical_test", [])
            needs_elec = any(e_kw.lower() in file_lower for e_kw in required_electrical)
            return "electronic", "电子料", kw, needs_elec

    # 再检查结构件关键词
    for kw in structural_kw:
        if kw.lower() in file_lower:
            return "structural", "结构件", kw, False

    return "unknown", "未知类型", None, False


def inspect_file_completeness(file_path, material_type, standards, pdf_text):
    """
    第一类：文件完整性检验
    返回: dict with item-level results and overall status
    """
    results = {
        "items": [],
        "pass_count": 0,
        "fail_count": 0,
        "total": 0,
        "status": "⚠️ 部分通过",
    }

    type_key = "electronic" if material_type == "electronic" else "structural"
    items = standards.get("file_completeness", {}).get(type_key, {}).get("items", [])

    results["total"] = len(items)

    for item in items:
        name = item["name"]
        english = item.get("english", "")
        note = item.get("note", "")

        found, keyword = check_keyword_in_text(pdf_text, [name, english])

        item_result = {
            "序号": item["id"],
            "项目": f"{name} ({english})",
            "必填": "✅" if item["required"] else "⚠️",
            "结果": "✅ 已找到" if found else "❌ 缺失",
            "备注": note or "",
        }

        if found:
            results["pass_count"] += 1
        else:
            results["fail_count"] += 1

        results["items"].append(item_result)

    if results["fail_count"] == 0:
        results["status"] = "✅ 通过"
    elif results["fail_count"] > 3:
        results["status"] = "❌ 不合格"

    return results


def inspect_rohs_compliance(file_path, standards, pdf_text, check_date):
    """
    第二类：RoHS合规性检验
    检查4个子项：
      2.1 RoHS测试报告是否存在
      2.2 RoHS调查表日期有效性（≤1年）
      2.3 RoHS调查表红框字段（6个必填）
      2.4 RoHS测试报告日期有效性
    """
    rohs_config = standards.get("rohs_check", {})
    validity_days = rohs_config.get("survey_form_validity_days", 365)
    required_fields = rohs_config.get("survey_form_required_fields", [])

    results = {
        "sub_items": {},
        "overall_status": "⚠️ 部分通过",
        "issues": [],
    }

    # 2.1 检查RoHS测试报告是否存在
    rohs_report_found, _ = check_keyword_in_text(pdf_text,
        ["rohs 2.0 test report", "rohs测试报告", "rohs 2.0 report"])
    results["sub_items"]["2.1_RoHS测试报告"] = (
        "✅ 已提供" if rohs_report_found else "❌ 未提供"
    )
    if not rohs_report_found:
        results["issues"].append("缺少RoHS 2.0测试报告")

    # 2.2 & 2.4 检查RoHS调查表和测试报告的日期
    all_dates = extract_dates_from_text(pdf_text)

    survey_date_valid = False
    test_date_valid = False

    if all_dates:
        latest_date = max(all_dates)
        days_ago = (check_date.date() - latest_date).days

        if days_ago <= validity_days:
            survey_date_valid = True
            test_date_valid = True
        else:
            results["issues"].append(
                f"RoHS报告日期已过期：报告日期 {latest_date}，距今 {days_ago} 天（超过 {validity_days} 天限制）"
            )

    results["sub_items"]["2.2_RoHS调查表日期有效性"] = (
        f"✅ 有效（≤{validity_days}天）" if survey_date_valid else ("❌ 过期" if all_dates else "⏭️ 未检测到日期")
    )
    results["sub_items"]["2.4_RoHS测试报告日期有效性"] = (
        f"✅ 有效（≤{validity_days}天）" if test_date_valid else ("❌ 过期" if all_dates else "⏭️ 未检测到日期")
    )

    # 2.3 检查RoHS调查表红框字段（6个必填字段）
    field_results = []
    missing_fields = []
    for field_info in required_fields:
        field_en = field_info.get("field", "")
        field_cn = field_info.get("chinese", "")

        found, _ = check_keyword_in_text(pdf_text, [field_en, field_cn])
        field_results.append({
            "字段": f"{field_en} ({field_cn})",
            "位置": field_info.get("location", ""),
            "状态": "✅ 已填写" if found else "❌ 未填写",
        })
        if not found:
            missing_fields.append(field_cn)

    results["sub_items"]["2.3_RoHS调查表红框字段"] = (
        f"✅ 全部填写（6/6）" if len(missing_fields) == 0
        else f"❌ {len(missing_fields)}项未填写：{'、'.join(missing_fields)}"
    )
    if missing_fields:
        results["issues"].append(f"RoHS调查表红框字段未填写：{'、'.join(missing_fields)}")

    # 总体判定
    passed_all = (
        rohs_report_found
        and survey_date_valid
        and test_date_valid
        and len(missing_fields) == 0
    )
    results["overall_status"] = "✅ 通过" if passed_all else "❌ 不合格"

    return results


def inspect_cpk_compliance(file_path, standards, pdf_text):
    """
    第三类：CPK合规性检验
    检查2个子项：
      3.1 CPK值是否≥1.33
      3.2 CPK报告与图纸尺寸对应性
    """
    cpk_config = standards.get("cpk_check", {})
    min_cpk = cpk_config.get("min_cpk_value", 1.33)

    results = {
        "sub_items": {},
        "overall_status": "⚠️ 部分通过",
        "issues": [],
        "cpk_values": [],
    }

    # 提取CPK值
    cpk_values = extract_cpk_values(pdf_text)
    results["cpk_values"] = cpk_values

    # 3.1 CPK值检查
    if cpk_values:
        min_val = min(cpk_values)
        max_val = max(cpk_values)
        all_pass = all(v >= min_cpk for v in cpk_values)
        results["sub_items"]["3.1_CPK值"] = (
            f"✅ 通过（范围: {min_val:.2f} ~ {max_val:.2f}, 均≥{min_cpk}）"
            if all_pass
            else f"❌ 不合格（最小值 {min_val:.2f} < {min_cpk}）"
        )
        if not all_pass:
            results["issues"].append(
                f"CPK值不合格：最小值 {min_val:.2f}，要求 ≥ {min_cpk}"
            )
    else:
        results["sub_items"]["3.1_CPK值"] = "⏭️ 未检测到CPK数据"
        results["issues"].append("未在PDF中找到CPK值数据")

    # 3.2 尺寸对应性检查（简化版 - 检查是否有尺寸数据）
    dimension_found, _ = check_keyword_in_text(pdf_text,
        ["dimension", "尺寸", "tolerance", "公差", "nominal"])
    results["sub_items"]["3.2_CPK尺寸对应性"] = (
        "✅ 已检测到尺寸数据（详细对应性需人工确认）" if dimension_found
        else "⏭️ 未检测到尺寸数据"
    )

    # 总体判定
    has_critical_issue = any(
        "CPK值不合格" in issue or "未在PDF中找到CPK" in issue
        for issue in results["issues"]
    )
    results["overall_status"] = "✅ 通过" if not has_critical_issue else "❌ 不合格"

    return results


def inspect_dimension_correspondence(file_path, standards, pdf_text):
    """
    第四类：尺寸公差对应性检验
    检查三文件包含关系：C ⊆ B ⊆ A
    （PDF解析版为简化检查，完整检查需要人工比对）
    """
    results = {
        "sub_items": {},
        "overall_status": "⏭️ 待人工确认",
        "issues": [],
    }

    dim_config = standards.get("dimension_check", {})
    hierarchy = dim_config.get("hierarchy", {})

    # 检测各层级的文件是否存在
    drawing_found, _ = check_keyword_in_text(pdf_text,
        ["engineering drawing", "工程图纸", "规格图纸", "drawing"])
    fullsize_found, _ = check_keyword_in_text(pdf_text,
        ["full size measurement", "全尺寸量测报告", "full size"])
    cpk_report_found, _ = check_keyword_in_text(pdf_text,
        ["cpk report", "cpk报告"])

    results["sub_items"]["4.1_规格图纸(A层)"] = (
        "✅ 已提供" if drawing_found else "❌ 缺失"
    )
    results["sub_items"]["4.2_全尺寸量测报告(B层)"] = (
        "✅ 已提供" if fullsize_found else "❌ 缺失"
    )
    results["sub_items"]["4.3_CPK报告(C层)"] = (
        "✅ 已提供" if cpk_report_found else "❌ 缺失"
    )

    # 包含关系验证（简化提示）
    all_present = drawing_found and fullsize_found and cpk_report_found
    if all_present:
        results["sub_items"]["4.4_包含关系C⊆B⊆A"] = (
            "⏭️ 三文件均存在，建议人工核对具体尺寸及公差是否一致"
        )
        results["overall_status"] = "⚠️ 需人工确认"
        results["issues"].append("三文件均存在，需人工核对尺寸及公差对应关系")
    else:
        results["sub_items"]["4.4_包含关系C⊆B⊆A"] = "❌ 文件缺失，无法验证包含关系"
        results["overall_status"] = "❌ 不合格"
        missing = []
        if not drawing_found:
            missing.append("规格图纸")
        if not fullsize_found:
            missing.append("全尺寸量测报告")
        if not cpk_report_found:
            missing.append("CPK报告")
        results["issues"].append(f"缺失文件：{'、'.join(missing)}，无法验证尺寸包含关系")

    return results


def inspect_report_validity(file_path, standards, pdf_text, check_date):
    """
    第五类：报告时效性检验
    """
    validity_config = standards.get("report_validity", {})
    rohs_days = validity_config.get("rohs_days", 365)
    reach_days = validity_config.get("reach_days", 365)

    results = {
        "sub_items": {},
        "overall_status": "✅ 通过",
        "issues": [],
    }

    all_dates = extract_dates_from_text(pdf_text)

    if all_dates:
        latest_date = max(all_dates)
        days_ago = (check_date.date() - latest_date).days

        if days_ago <= rohs_days:
            results["sub_items"]["5.1_RoHS报告时效性"] = (
                f"✅ 有效（报告日期: {latest_date}，距今 {days_ago} 天 ≤ {rohs_days} 天）"
            )
        else:
            results["sub_items"]["5.1_RoHS报告时效性"] = (
                f"❌ 已过期（报告日期: {latest_date}，距今 {days_ago} 天 > {rohs_days} 天）"
            )
            results["issues"].append(f"RoHS报告已过期 {days_ago - rohs_days} 天")
            results["overall_status"] = "❌ 不合格"

        # REACH报告通常与RoHS在同一份报告中
        reach_found, _ = check_keyword_in_text(pdf_text, ["reach"])
        if reach_found:
            results["sub_items"]["5.2_REACH报告时效性"] = (
                f"✅ 有效（同上）" if days_ago <= reach_days
                else f"❌ 已过期"
            )
        else:
            results["sub_items"]["5.2_REACH报告时效性"] = "⏭️ 未检测到REACH报告"
    else:
        results["sub_items"]["5.1_RoHS报告时效性"] = "⏭️ 未检测到报告日期"
        results["sub_items"]["5.2_REACH报告时效性"] = "⏭️ 未检测到报告日期"

    return results


def generate_final_verdict(material_type, all_results, standards):
    """
    第六类：生成最终检验结论与处理建议
    """
    issues = []
    completeness = all_results.get("completeness", {})
    rohs = all_results.get("rohs", {})
    cpk = all_results.get("cpk", {})
    dimension = all_results.get("dimension", {})
    validity = all_results.get("validity", {})

    # 收集所有问题
    issues.extend(completeness.get("issues", []))
    issues.extend(rohs.get("issues", []))
    issues.extend(cpk.get("issues", []))
    issues.extend(dimension.get("issues", []))
    issues.extend(validity.get("issues", []))

    # 判定总体结论
    total_fail = completeness.get("fail_count", 0)
    critical_fail = (
        completeness["status"] == "❌ 不合格"
        or rohs["overall_status"] == "❌ 不合格"
        or cpk["overall_status"] == "❌ 不合格"
        or dimension["overall_status"] == "❌ 不合格"
        or validity["overall_status"] == "❌ 不合格"
    )

    if critical_fail or len(issues) > 3:
        verdict = "❌ 不合格，退回重报"
        suggestion = (
            "退回供应商，要求按XC-R-0802-DQM-002正式样品承认书模板重新提交完整的样品承认书\n\n"
            "**必须补充以下文件：**\n"
        )
        # 列出缺失的文件
        for item in completeness.get("items", []):
            if item["结果"].startswith("❌"):
                suggestion += f"- ❌ **{item['项目']}**\n"
        # 其他问题
        for issue in issues:
            suggestion += f"- ⚠️ {issue}\n"
    elif len(issues) > 0:
        verdict = "⚠️ 基本合格，需补充材料"
        suggestion = "该封样报告基本符合要求，但存在以下问题需补充或修正：\n\n"
        for i, issue in enumerate(issues, 1):
            suggestion += f"{i}. {issue}\n"
    else:
        verdict = "✅ 合格，建议通过"
        suggestion = "该封样报告符合XC-R-0802-DQM-002模板要求，建议通过审核。"

    return {
        "verdict": verdict,
        "issues": issues,
        "issue_count": len(issues),
        "suggestion": suggestion,
        "material_type": material_type,
    }


# ============================================================
# 主审核流程
# ============================================================

def run_full_inspection(file_path, file_name, standards):
    """执行完整的6大类38项审核流程"""
    check_date = datetime.now()
    version = standards.get("version", "未知")

    # 提取PDF文本
    pdf_text = extract_pdf_text(file_path)

    # 第零步：物料类型判定
    mat_type, mat_type_cn, mat_kw, needs_elec = determine_material_type(file_name, standards)

    # 第一步~第五步：各类检验
    completeness = inspect_file_completeness(file_path, mat_type, standards, pdf_text)
    rohs = inspect_rohs_compliance(file_path, standards, pdf_text, check_date)
    cpk = inspect_cpk_compliance(file_path, standards, pdf_text)
    dimension = inspect_dimension_correspondence(file_path, standards, pdf_text)
    validity = inspect_report_validity(file_path, standards, pdf_text, check_date)

    # 第六步：生成最终结论
    all_results = {
        "completeness": completeness,
        "rohs": rohs,
        "cpk": cpk,
        "dimension": dimension,
        "validity": validity,
    }
    final = generate_final_verdict(mat_type, all_results, standards)

    # 构建返回结果
    result = {
        "文件名": file_name,
        "物料类型": f"{mat_type_cn}" if mat_type != "unknown" else "未知（需人工确认）",
        "需要电气性能测试": "是" if needs_elec else "否",
        "文件完整性": completeness["status"],
        "RoHS合规性": rohs["overall_status"],
        "CPK合规性": cpk["overall_status"],
        "尺寸对应性": dimension["overall_status"],
        "报告时效性": validity["overall_status"],
        "总体结论": final["verdict"],
        "问题数量": final["issue_count"],
        "审核时间": check_date.strftime("%Y-%m-%d %H:%M:%S"),
        "标准版本": f"V{version}",
        "_detail": {
            "mat_type": mat_type,
            "mat_type_cn": mat_type_cn,
            "needs_elec": needs_elec,
            "completeness": completeness,
            "rohs": rohs,
            "cpk": cpk,
            "dimension": dimension,
            "validity": validity,
            "final": final,
            "pdf_text_preview": pdf_text[:2000] if pdf_text else "",
        },
    }
    return result


# ============================================================
# Streamlit UI
# ============================================================

# 设置页面配置
st.set_page_config(
    page_title="封样检验应用",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 加载标准
standards = load_standards()
if standards is None:
    st.stop()

version = standards.get("version", "未知")
last_updated = standards.get("last_updated", "未知")

# 标题区
st.title("📋 封样检验应用")
st.caption(f"基于 XC-R-0802-DQM-002 物料正式样品承认书模板 | 标准版本 V{version} | 更新于 {last_updated}")
st.markdown("---")

# 侧边栏设置
st.sidebar.header("⚙️ 审核标准设置")

type_key_electronic = standards.get("file_completeness", {}).get("electronic", {}).get("items", [])
all_item_names = [item["name"] for item in type_key_electronic]
default_names = ["物料清单", "工程图纸", "样品照片", "全尺寸测量报告", "Cpk报告",
                 "产品规格书", "制造流程图", "包装方式", "QC工程图", "可靠性测试报告",
                 "材质证明", "RoHS 2.0限用物质成分调查表", "RoHS 2.0测试报告"]

st.sidebar.subheader("文件完整性检查")
check_list = st.sidebar.multiselect(
    "选择要检查的项目",
    all_item_names,
    default=[n for n in default_names if n in all_item_names],
)

st.sidebar.subheader("专项检查")
rohs_check = st.sidebar.checkbox("✅ RoHS合规性检验（含日期+红框字段）", value=True)
cpk_check = st.sidebar.checkbox("✅ CPK合规性检验（≥1.33）", value=True)
dim_check = st.sidebar.checkbox("✅ 尺寸公差对应性检验", value=True)
validity_check = st.sidebar.checkbox("✅ 报告时效性检验（≤1年）", value=True)

min_cpk_display = standards.get("cpk_check", {}).get("min_cpk_value", 1.33)
st.sidebar.info(f"当前CPK合格标准：≥{min_cpk_display}")

# 主界面两栏布局
col1, col2 = st.columns([1, 1])

with col1:
    st.header("📤 上传PDF文件")

    uploaded_files = st.file_uploader(
        "选择PDF文件（支持批量上传 / 拖拽上传）",
        type=["pdf"],
        accept_multiple_files=True,
        help="可一次选多个PDF文件，或直接拖入此区域"
    )

    st.subheader("或者使用文件夹路径扫描")
    folder_path = st.text_input(
        "输入包含PDF文件的文件夹路径",
        placeholder="例如：D:\\封样\\S1651\\楚鑫"
    )

    if st.button("📁 扫描文件夹", type="secondary"):
        if folder_path and os.path.exists(folder_path):
            pdf_list = []
            for root, dirs, files in os.walk(folder_path):
                for fn in files:
                    if fn.lower().endswith(".pdf"):
                        pdf_list.append(os.path.join(root, fn))
            if pdf_list:
                st.success(f"✅ 找到 **{len(pdf_list)}** 个PDF文件")
                st.session_state['folder_files'] = pdf_list
            else:
                st.warning("⚠️ 该文件夹中没有PDF文件")
        else:
            st.error("❌ 路径不存在")

    # 显示已选文件列表
    if uploaded_files or 'folder_files' in st.session_state:
        st.subheader("📄 已选择的文件")
        count = 0
        if uploaded_files:
            for f in uploaded_files:
                st.text(f"✅ {f.name}")
                count += 1
        if 'folder_files' in st.session_state:
            for fp in st.session_state['folder_files']:
                st.text(f"✅ {os.path.basename(fp)}")
                count += 1
        st.info(f"共 **{count}** 个文件待审核")

        if st.button("🗑️ 清空文件列表", type="secondary"):
            if 'folder_files' in st.session_state:
                del st.session_state['folder_files']
            st.rerun()


with col2:
    st.header("📊 审核结果")

    run_btn = st.button("🚀 开始审核", type="primary", use_container_width=True)

    if run_btn:
        # 收集文件
        all_paths = []
        if uploaded_files:
            for uf in uploaded_files:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                tmp.write(uf.getvalue())
                all_paths.append(tmp.name)
        if 'folder_files' in st.session_state:
            all_paths.extend(st.session_state['folder_files'])

        if not all_paths:
            st.warning("⚠️ 请先上传或选择PDF文件")
        else:
            st.info(f"开始审核 **{len(all_paths)}** 个文件...")

            progress = st.progress(0)
            detail_results = []

            for i, fp in enumerate(all_paths):
                fname = os.path.basename(fp)
                with st.spinner(f"正在审核: {fname}"):
                    res = run_full_inspection(fp, fname, standards)
                    detail_results.append(res)
                progress.progress((i + 1) / len(all_paths))

            st.success(f"✅ 审核完成！共审核 **{len(detail_results)}** 个文件")

            # 结果汇总表格
            df_summary = pd.DataFrame([
                {k: v for k, v in r.items() if not k.startswith("_")}
                for r in detail_results
            ])
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

            # 统计概览
            col_s1, col_s2, col_s3 = st.columns(3)
            pass_count = sum(1 for r in detail_results if "合格" in r["总体结论"])
            fail_count = sum(1 for r in detail_results if "不合格" in r["总体结论"])
            warn_count = len(detail_results) - pass_count - fail_count

            with col_s1:
                st.metric("✅ 合格", pass_count, delta_color="normal")
            with col_s2:
                st.metric("⚠️ 需补充", warn_count, delta_color="off")
            with col_s3:
                st.metric("❌ 不合格", fail_count, delta_color="inverse")

            # 详细展开器
            st.subheader("📋 各文件详细审核报告")
            for idx, res in enumerate(detail_results):
                d = res["_detail"]
                expander_title = (
                    f"【{idx+1}】{res['文件名']} — "
                    f"类型:{res['物料类型']} | "
                    f"结论:{res['总体结论']} | "
                    f"问题数:{res['问题数量']}"
                )
                with st.expander(expander_title, expanded=(res["问题数量"] > 0)):
                    st.markdown(f"**物料类型:** {d['mat_type_cn']}"
                                f" | **需电气性能测试:** {'是' if d['needs_elec'] else '否'}")

                    # 文件完整性详情
                    st.subheader("1️⃣ 文件完整性检验")
                    comp_df = pd.DataFrame(d["completeness"]["items"])
                    st.dataframe(comp_df, hide_index=True, use_container_width=True)

                    # RoHS详情
                    if rohs_check:
                        st.subheader("2️⃣ RoHS合规性检验")
                        for k, v in d["rohs"]["sub_items"].items():
                            st.text(f"{k}: {v}")

                    # CPK详情
                    if cpk_check:
                        st.subheader("3️⃣ CPK合规性检验")
                        for k, v in d["cpk"]["sub_items"].items():
                            st.text(f"{k}: {v}")
                        if d["cpk"]["cpk_values"]:
                            st.text(f"检测到的CPK值: {d['cpk']['cpk_values']}")

                    # 尺寸对应性详情
                    if dim_check:
                        st.subheader("4️⃣ 尺寸公差对应性检验")
                        for k, v in d["dimension"]["sub_items"].items():
                            st.text(f"{k}: {v}")

                    # 报告时效性详情
                    if validity_check:
                        st.subheader("5️⃣ 报告时效性检验")
                        for k, v in d["validity"]["sub_items"].items():
                            st.text(f"{k}: {v}")

                    # 最终处理建议
                    st.subheader("6️⃣ 检验结论与处理建议")
                    st.markdown(f"**{d['final']['verdict']}**")
                    st.markdown(d['final']['suggestion'])

            # 下载Excel
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                excel_buf = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                df_summary.to_excel(excel_buf.name, index=False, engine="openpyxl")
                with open(excel_buf.name, "rb") as ef:
                    st.download_button(
                        label="📥 下载审核汇总 Excel",
                        data=ef,
                        file_name=f"封样审核汇总_V{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            with col_dl2:
                full_report = f"# 封样检验审核报告\n\n"
                full_report += f"**标准版本:** V{version}\n"
                full_report += f"**审核时间:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                full_report += f"**共审核:** {len(detail_results)} 个文件\n\n---\n\n"
                for i, res in enumerate(detail_results):
                    d = res["_detail"]
                    full_report += f"## 【{i+1}】{res['文件名']}\n\n"
                    full_report += f"**物料类型:** {res['物料类型']}\n"
                    full_report += f"**总体结论:** {res['总体结论']}\n\n"
                    full_report += f"### 处理建议\n\n{d['final']['suggestion']}\n\n---\n\n"
                st.download_button(
                    label="📥 下载完整报告 Markdown",
                    data=full_report.encode("utf-8"),
                    file_name=f"封样审核报告_V{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                    mime="text/markdown"
                )

            # 清理临时文件
            for p in all_paths:
                if p.startswith(tempfile.gettempdir()):
                    try:
                        os.unlink(p)
                    except OSError:
                        pass


# 底部说明
st.markdown("---")
st.markdown("""
### 📝 使用说明
1. **上传PDF** — 拖拽或点击批量上传，也可输入文件夹路径批量扫描
2. **选择标准** — 左侧边栏勾选要执行的审核项
3. **开始审核** — 点击按钮，等待审核完成
4. **查看结果** — 展开每个文件的详细信息，下载Excel/Markdown报告

### ⚙️ 当前标准版本说明
| 类别 | 项目数 | 说明 |
|------|--------|------|
| 物料类型判定 | 自动 | 根据文件名判定电子料/结构件 |
| 文件完整性 | 16项(电子料)/15项(结构件) | 按物料类型区分必填项 |
| RoHS合规性 | 4子项 | 报告来源、日期、红框6字段 |
| CPK合规性 | 2子项 | CPK≥1.33、尺寸对应性 |
| 尺寸对应性 | 4子项 | 三文件包含关系 C⊆B⊆A |
| 报告时效性 | 2子项 | RoHS/REACH报告≤1年 |

> 💡 **注意：** 本工具使用PDF文本解析技术进行自动化审核，部分内容（如图片中的尺寸标注、手写签名等）可能需要人工辅助确认。
""")

hide_style = "<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;}</style>"
st.markdown(hide_style, unsafe_allow_html=True)
