#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
封样检验Web应用 - V6.2 精准审核版
基于 SKILL.md V4.0 (2026-06-23)
实现PDF逐页分析、工程图纸判定规则、产品规格书判定规则
V6.2新增：目录勾选状态检测、料号&物料名称跨表一致性检查
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
# V4.0 新增：PDF逐页分析引擎
# ============================================================

def analyze_pdf_page_by_page(pdf_path):
    """
    V4.0 步骤2：逐页分析PDF内容类型
    返回：list of dict, 每个dict代表一页的分析结果
    """
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_num = page.page_number
                text = page.extract_text() or ""
                text_lower = text.lower()

                page_info = {
                    "page_num": page_num,
                    "text": text,
                    "text_lower": text_lower,
                    "is_cover": False,
                    "is_engineering_drawing": False,
                    "is_bom": False,
                    "is_sample_photo": False,
                    "is_cpk": False,
                    "is_rohs_survey": False,
                    "is_reach_survey": False,
                    "is_product_spec": False,
                    "drawing_number": "",
                    "drawing_version": "",
                    "dimensions": [],
                }

                # --- 封面页判定 ---
                if page_num <= 3:
                    if any(kw in text_lower for kw in ["provisional sample", "临时样品承认书", "sample acknowledgement"]):
                        page_info["is_cover"] = True

                # --- 工程图纸页判定（V4.0 步骤3规则）---
                drawing_score = 0
                # 规则1：含图号（如 NC-XC260522-007）
                drawing_match = re.search(r'[A-Za-z]{2,}-[A-Za-z0-9\-]{3,}', text)
                if drawing_match:
                    page_info["drawing_number"] = drawing_match.group(0)
                    drawing_score += 3
                # 规则2：含版本号
                version_match = re.search(r'(版本|rev)\s*[：:]?\s*[A-Za-z0-9]', text_lower)
                if version_match:
                    page_info["drawing_version"] = version_match.group(0)
                    drawing_score += 2
                # 规则3：含尺寸标注（如 2000±50、5.4±0.2、OD、ID）
                dim_pattern = r'\d{2,}\s*[±±]\s*\d+|\b(od|id|thk|len)\s*[:：]?\s*\d'
                if re.search(dim_pattern, text_lower):
                    dims = re.findall(r'\d+\.?\d*\s*[±±]\s*\d+', text)
                    page_info["dimensions"] = dims[:10]
                    drawing_score += 3
                # 规则4：含标题栏（含 批准/审核/制定）
                if any(kw in text for kw in ["批准", "审核", "制定", "设计", "approve", "check"]):
                    drawing_score += 1
                # 规则5：含比例、视角
                if any(kw in text_lower for kw in ["比例", "scale", "视角", "view", "top view", "side view"]):
                    drawing_score += 1

                if drawing_score >= 4:
                    page_info["is_engineering_drawing"] = True

                # --- BOM表页判定 ---
                if "bom" in text_lower or "物料清单" in text or "bill of material" in text_lower:
                    page_info["is_bom"] = True

                # --- 样品照片页判定 ---
                if "附上样品实物" in text or "sample photo" in text_lower or "样品照片" in text:
                    page_info["is_sample_photo"] = True

                # --- CPK报告页判定（V6.1 增强：支持表格形式/繁体中文/多种格式）---
                cpk_indicators = [
                    r'cp[kK]\s*[:：=]',                    # cpk: / cpk= / cpk：
                    r'cp[kK]\s*報告',                      # 繁体 CPK報告
                    r'cp[kK]\s*报告',                      # 简体 CPK报告
                    r'cp[kK]\s*report',                   # CPK report
                    r'cp[kK]\s*值',                        # CPK值
                    r'cp[kK]\s*[\(（].*?[\)）]',          # CPK(xxx)
                    r'^.*cp[kK].*$',                       # 任意含CPK的行（宽泛匹配）
                ]
                # 宽松匹配：页面文本中只要包含 CPK/Cpk 关键词即判定
                if re.search(r'cp[kK]', text_lower):
                    # 进一步确认不是偶然出现（至少出现2次或有相关统计关键词）
                    cpk_count = len(re.findall(r'cp[kK]', text_lower))
                    has_cpk_context = any(kw in text_lower for kw in [
                        'usl', 'lsl', 'ppk', 'stddev', 'average', 'range',
                        '规格上限', '规格下限', '标准差', '平均值', '量測', '测量',
                        'dimension', 'tolerance', 'nominal', 'specification'
                    ])
                    if cpk_count >= 2 or has_cpk_context:
                        page_info["is_cpk"] = True

                # --- RoHS调查表页判定 ---
                if "rohs" in text_lower and ("调查表" in text or "survey" in text_lower):
                    page_info["is_rohs_survey"] = True

                # --- REACH调查表页判定 ---
                if "reach" in text_lower and ("调查表" in text or "survey" in text_lower):
                    page_info["is_reach_survey"] = True

                # --- 产品规格书/技术规范判定（V4.0 步骤4规则）---
                spec_score = 0
                if any(kw in text for kw in ["绝缘电阻", "insulation resistance", "Ω"]):
                    spec_score += 2
                if any(kw in text for kw in ["PVC", "材质", "material spec", "执行标准", "standard"]):
                    spec_score += 1
                if any(kw in text_lower for kw in ["100%电气测试", "electrical test", "eia/tia"]):
                    spec_score += 2
                if spec_score >= 2:
                    page_info["is_product_spec"] = True

                results.append(page_info)

    except Exception as e:
        results.append({"error": str(e)})

    return results


def check_file_type(pdf_path, pdf_text):
    """
    V4.0 第一步：文件类型检查
    返回: (file_type, note)
    file_type: "DQM-002" / "DQM-001" / "unknown"
    """
    text_lower = pdf_text.lower()
    if "xc-r-0802-dqm-002" in text_lower or "正式样品承认书" in text_lower:
        return "DQM-002", "正式样品承认书"
    if "xc-r-0802-dqm-001" in text_lower or "临时样品承认书" in text_lower:
        return "DQM-001", "⚠️ 临时样品承认书（与正式模板不符）"
    # 从文件名判断
    fname = os.path.basename(pdf_path).upper()
    if "DQM-002" in fname:
        return "DQM-002", "从文件名判定为正式样品承认书"
    if "DQM-001" in fname:
        return "DQM-001", "⚠️ 从文件名判定为临时样品承认书"
    return "unknown", "未识别文件类型编号"


def check_engineering_drawing_detailed(page_analysis):
    """
    V4.0 步骤3：工程图纸详细判定
    返回: (is_present, details_dict)
    """
    has_drawing = any(p.get("is_engineering_drawing") for p in page_analysis)
    details = {
        "has_drawing": has_drawing,
        "drawing_pages": [p["page_num"] for p in page_analysis if p.get("is_engineering_drawing")],
        "drawing_numbers": list(set(p.get("drawing_number", "") for p in page_analysis if p.get("drawing_number"))),
        "has_dimensions": any(p.get("dimensions") for p in page_analysis),
        "note": "",
    }
    if has_drawing:
        details["note"] = f"✅ 已提供工程图纸（第 {', '.join(map(str, details['drawing_pages']))} 页）"
        if details["drawing_numbers"]:
            details["note"] += f"，图号：{', '.join(details['drawing_numbers'])}"
    else:
        # 即使未标记为工程图纸，检查是否有含图号的页面
        has_dn = any(p.get("drawing_number") for p in page_analysis)
        if has_dn:
            details["note"] = "⚠️ 页面含图号但未标记为工程图纸，建议人工确认"
            details["has_drawing"] = True  # 保守判定为已提供
        else:
            details["note"] = "❌ 未检测到工程图纸页"
    return details["has_drawing"], details


def check_product_specification(page_analysis):
    """
    V4.0 步骤4：产品规格书判定
    返回: (status, note)
    status: "独立提供" / "部分提供（嵌入工程图纸）" / "缺失"
    """
    # 检查是否有独立的产品规格书页
    has_spec_page = any(p.get("is_product_spec") for p in page_analysis)
    # 检查工程图纸页是否含技术规格内容
    drawing_pages = [p for p in page_analysis if p.get("is_engineering_drawing") or p.get("drawing_number")]
    has_spec_in_drawing = False
    if drawing_pages:
        for p in drawing_pages:
            txt = p.get("text", "")
            if any(kw in txt for kw in ["绝缘电阻", "PVC", "执行标准", "Ω", "材质", "100%电气"]):
                has_spec_in_drawing = True
                break

    if has_spec_page:
        return "独立提供", "✅ 已提供独立的产品规格书"
    elif has_spec_in_drawing:
        return "部分提供", "⚠️ 技术要求嵌入在工程图纸中，建议补充独立的产品规格书文档"
    else:
        return "缺失", "❌ 未检测到产品规格书内容"


# ============================================================
# PDF解析引擎（基础）
# ============================================================

def extract_all_text(pdf_path):
    """提取PDF全部文本（合并）"""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
    except Exception as e:
        text = f"[PDF解析错误: {e}]"
    return text


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
    patterns = [
        r'(\d{4})[/\-\.](\d{1,2})[/\-\.](\d{1,2})',
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


def extract_cpk_values(text, pdf_path=None):
    """
    V6.1 增强：从文本和表格中提取CPK值
    支持多种格式：
    1. 文本键值对：cpk: 1.33 / cpk=1.33
    2. 表格形式：CPK列头 + 数值行（pdfplumber extract_tables）
    3. 统计行格式：... | CPK | 0.67 | 3.84 | ...
    4. 宽松相邻匹配：CPK/Cpk 紧邻数字
    """
    cpk_values = []
    seen = set()

    def add_val(v):
        """去重添加CPK值"""
        if v not in seen and 0 <= v <= 100:  # 合理的CPK范围
            seen.add(v)
            cpk_values.append(v)

    # ===== 方法1：文本正则匹配（原有逻辑 + 扩展）=====
    text_patterns = [
        r'cpk\s*[:：=]\s*(\d+\.?\d*)',           # cpk: 1.33
        r'cpk\s*[\(（]\s*(\d+\.?\d*)',           # cpk(1.33)
        r'cpk\s+value\s*[:：=]\s*(\d+\.?\d*)',   # cpk value: 1.33
        r'cp\s*[kK]\s*[:：=]\s*(\d+\.?\d*)',     # CP : 1.33
        r'c pk\s*[:：=]\s*(\d+\.?\d*)',          # c pk: 1.33 (可能的空格)
        # V6.1 新增：宽松相邻匹配
        r'cp[kK]\s*[\/>]?\s*(\d+\.\d{1,2})',     # CPK>1.33 / CPK/1.33 / CPK 1.33
        r'cp[kK]\s+(?:最小|min|minimum)?\s*[:：]?\s*(\d+\.\d+)',  # CPK最小: 1.33
        r'(?:cpk|Cpk|CPK)\s.*?(\d+\.\d{2})',      # CPK xxx 1.33（同一行内）
    ]
    for pattern in text_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for val in matches:
            try:
                add_val(float(val))
            except ValueError:
                continue

    # ===== 方法2：从PDF表格中提取CPK值（V6.1 核心新增）=====
    if pdf_path:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if not tables:
                        continue
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        # 寻找包含 CPK 的表头行
                        for row_idx, row in enumerate(table):
                            if row is None:
                                continue
                            # 将row中所有单元格转为字符串
                            row_strs = [str(c).strip() if c else "" for c in row]
                            row_text_lower = " ".join(row_strs).lower()

                            # 检查此行是否含CPK相关表头
                            if re.search(r'cp[kK]', row_text_lower):
                                # 找到CPK列的索引
                                cpk_col_indices = []
                                for col_idx, cell in enumerate(row_strs):
                                    if re.search(r'cp[kK]', cell.lower()):
                                        cpk_col_indices.append(col_idx)

                                # 如果找到了CPK列，从后续数据行中取值
                                if cpk_col_indices:
                                    # 查看后续数据行（最多往后查10行）
                                    for data_row_idx in range(row_idx + 1, min(row_idx + 11, len(table))):
                                        data_row = table[data_row_idx]
                                        if data_row is None:
                                            continue
                                        for col_idx in cpk_col_indices:
                                            cell_val = data_row[col_idx] if col_idx < len(data_row) else None
                                            if cell_val is not None:
                                                val_str = str(cell_val).strip()
                                                # 提取数字
                                                num_match = re.match(r'^(\d+\.?\d*)$', val_str)
                                                if num_match:
                                                    try:
                                                        add_val(float(num_match.group(1)))
                                                    except ValueError:
                                                        pass
                                                # 也尝试更宽松的提取
                                                elif re.search(r'\d+\.\d{1,2}', val_str):
                                                    nums = re.findall(r'\d+\.\d{1,2}', val_str)
                                                    for n in nums:
                                                        try:
                                                            add_val(float(n))
                                                        except ValueError:
                                                            pass
                                    break  # 只处理第一个含CPK的表头行

                        # 额外：扫描整个表格中的CPK数值模式
                        # 有些PDF的表格结构不规则，直接搜索所有单元格
                        for row in table:
                            if row is None:
                                continue
                            for cell in row:
                                if cell is None:
                                    continue
                                cell_str = str(cell).strip()
                                # 匹配如 "CPK=0.67" 或 "Cpk: 3.84" 格式的单元格
                                m = re.search(r'cp[kK]\s*[:：=]?\s*(\d+\.?\d*)', cell_str, re.IGNORECASE)
                                if m:
                                    try:
                                        add_val(float(m.group(1)))
                                    except ValueError:
                                        pass
        except Exception:
            pass  # 表格提取失败时静默跳过，不影响主流程

    # ===== 方法3：多行统计摘要匹配 =====
    # 匹配类似 "CPK  0.67  3.84  4.00" 这种同行多值格式
    multi_cpk_pattern = r'cp[kK][\s\:：]*([0-9.]+(?:\s+[0-9.]+)*)'
    multi_matches = re.findall(multi_cpk_pattern, text, re.IGNORECASE)
    for match_group in multi_matches:
        numbers = re.findall(r'(\d+\.\d{1,2})', match_group)
        for n in numbers:
            try:
                v = float(n)
                # 过滤掉明显不是CPK值的数（如尺寸值>100或过小的噪声值<0.01）
                if 0.01 <= v <= 20:
                    add_val(v)
            except ValueError:
                continue

    return cpk_values


def check_keyword_in_text(text, keywords):
    """检查文本中是否包含关键词列表中的任一关键词"""
    for kw in keywords:
        if kw.lower() in text.lower():
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


def inspect_file_completeness_v4(page_analysis, material_type, standards):
    """
    V4.0 第一类：文件完整性检验（基于逐页分析结果）
    返回: dict with item-level results and overall status
    """
    results = {
        "items": [],
        "pass_count": 0,
        "fail_count": 0,
        "total": 0,
        "status": "⚠️ 部分通过",
        "page_analysis_summary": {},
    }

    type_key = "electronic" if material_type == "electronic" else "structural"
    items = standards.get("file_completeness", {}).get(type_key, {}).get("items", [])

    results["total"] = len(items)

    # 基于逐页分析结果判断
    has_drawing = any(p.get("is_engineering_drawing") or p.get("drawing_number") for p in page_analysis)
    has_sample_photo = any(p.get("is_sample_photo") for p in page_analysis)
    has_cpk = any(p.get("is_cpk") for p in page_analysis)
    has_bom = any(p.get("is_bom") for p in page_analysis)
    has_rohs = any(p.get("is_rohs_survey") for p in page_analysis)
    has_reach = any(p.get("is_reach_survey") for p in page_analysis)
    spec_status, spec_note = check_product_specification(page_analysis)

    # 全文本用于其他检查
    all_text = " ".join(p.get("text", "") for p in page_analysis).lower()

    for item in items:
        name = item["name"]
        english = item.get("english", "")
        note = item.get("note", "")
        required = item.get("required", True)

        found = False
        item_note = note or ""

        # 基于页面分析结果判断
        if "物料清单" in name or "Bill of Material" in english:
            found = has_bom
        elif "工程图纸" in name or "Engineering Drawing" in english:
            found, draw_details = check_engineering_drawing_detailed(page_analysis)
            item_note = draw_details.get("note", item_note)
        elif "样品照片" in name or "Sample Photos" in english:
            found = has_sample_photo
        elif "全尺寸" in name or "Full Size" in english:
            # 全尺寸报告通常含尺寸测量表格
            found = "全尺寸" in all_text or "full size" in all_text or "measurement report" in all_text
        elif "Cpk" in name or "Cpk Report" in english:
            found = has_cpk
        elif "产品规格书" in name or "Product Specification" in english:
            found = spec_status != "缺失"
            item_note = spec_note
        elif "制造流程图" in name or "Process Flow" in english:
            found = "制造流程" in all_text or "process flow" in all_text
        elif "包装方式" in name or "Packaging" in english:
            found = "包装" in all_text or "packaging" in all_text
        elif "QC工程图" in name or "QC Flow" in english:
            found = "qc flow" in all_text or "qc 工程" in all_text
        elif "电气性能" in name or "Electrical Performance" in english:
            found = "电气性能" in all_text or "electrical performance" in all_text
        elif "可靠性" in name or "Reliability" in english:
            found = "可靠性" in all_text or "reliability" in all_text
        elif "材质证明" in name or "Material Certificate" in english:
            found = "材质证明" in all_text or "material certificate" in all_text or "sgs" in all_text
        elif "RoHS" in name and "调查表" in name:
            found = has_rohs
        elif "RoHS" in name and "测试报告" in name:
            found = "rohs" in all_text and "test report" in all_text
        elif "REACH" in name and "调查表" in name:
            found = has_reach
        elif "REACH" in name and "测试报告" in name:
            found = "reach" in all_text and "test report" in all_text
        else:
            # 通用关键词匹配
            fnd, _ = check_keyword_in_text(all_text, [name, english])
            found = fnd

        item_result = {
            "序号": item["id"],
            "项目": f"{name} ({english})",
            "必填": "✅" if required is True else ("⚠️" if required == "conditional" else "❌"),
            "结果": "✅ 已找到" if found else "❌ 缺失",
            "备注": item_note,
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


def inspect_rohs_compliance(page_analysis, standards, check_date):
    """
    第二类：RoHS合规性检验（4子项）
    """
    results = {
        "sub_items": {},
        "overall_status": "⚠️ 部分通过",
        "issues": [],
    }

    all_text = " ".join(p.get("text", "") for p in page_analysis)
    all_text_lower = all_text.lower()

    # 2.1 RoHS 2.0测试报告是否存在
    rohs_report_found = (
        "rohs 2.0 test report" in all_text_lower or
        "rohs测试报告" in all_text_lower or
        "rohs 2.0 report" in all_text_lower
    )
    results["sub_items"]["2.1_RoHS测试报告"] = (
        "✅ 已提供" if rohs_report_found else "❌ 未提供"
    )
    if not rohs_report_found:
        results["issues"].append("缺少RoHS 2.0测试报告")

    # 2.2 & 2.4 报告日期有效性
    all_dates = extract_dates_from_text(all_text)
    survey_date_valid = False
    test_date_valid = False

    if all_dates:
        latest_date = max(all_dates)
        days_ago = (check_date.date() - latest_date).days
        if days_ago <= 365:
            survey_date_valid = True
            test_date_valid = True
        else:
            results["issues"].append(
                f"RoHS报告日期过期：报告日期 {latest_date}，距今 {days_ago} 天（超过365天限制）"
            )

    results["sub_items"]["2.2_RoHS调查表日期有效性"] = (
        f"✅ 有效（≤365天）" if survey_date_valid else ("❌ 过期" if all_dates else "⏱️ 未检测到日期")
    )
    results["sub_items"]["2.4_RoHS测试报告日期有效性"] = (
        f"✅ 有效（≤365天）" if test_date_valid else ("❌ 过期" if all_dates else "⏱ 未检测到日期")
    )

    # 2.3 RoHS调查表红框字段（6个必填字段）
    # 检查页面中含RoHS调查表的页面
    rohs_survey_pages = [p for p in page_analysis if p.get("is_rohs_survey")]
    red_box_fields = [
        ("Monomer（单体）", ["monomer", "单体"]),
        ("Supplier（供应商）", ["supplier", "供应商"]),
        ("Control method（控制方法）", ["control method", "控制方法"]),
        ("Number（编号）", ["number", "编号", "report number"]),
        ("Effective date of report（出报告日期）", ["effective date", "出报告日期", "报告日期"]),
        ("Remarks（备注）", ["remarks", "备注"]),
    ]
    missing_fields = []
    for field_name, keywords in red_box_fields:
        found = False
        for p in rohs_survey_pages:
            txt = p.get("text", "")
            if any(kw.lower() in txt.lower() for kw in keywords):
                found = True
                break
        if not found:
            # 在全文中再查一次
            if not any(kw.lower() in all_text_lower for kw in keywords):
                missing_fields.append(field_name.split("（")[0])

    fill_status = f"✅ 全部填写（{6 - len(missing_fields)}/6）" if len(missing_fields) == 0 else f"❌ {len(missing_fields)}项未填写：{'、'.join(missing_fields)}"
    results["sub_items"]["2.3_RoHS调查表红框字段"] = fill_status
    if missing_fields:
        results["issues"].append(f"RoHS调查表红框字段未填写：{'、'.join(missing_fields)}")

    passed_all = rohs_report_found and survey_date_valid and test_date_valid and len(missing_fields) == 0
    results["overall_status"] = "✅ 通过" if passed_all else "❌ 不合格"

    return results


def inspect_cpk_compliance(page_analysis, standards, pdf_text, pdf_path=None):
    """
    第三类：CPK合规性检验（2子项）
    V6.1: 支持从PDF表格中提取CPK值
    """
    results = {
        "sub_items": {},
        "overall_status": "⚠️ 部分通过",
        "issues": [],
        "cpk_values": [],
    }

    all_text = " ".join(p.get("text", "") for p in page_analysis)

    # 3.1 提取CPK值（V6.1: 传入pdf_path以支持表格解析）
    cpk_values = extract_cpk_values(all_text, pdf_path=pdf_path)
    results["cpk_values"] = cpk_values

    if cpk_values:
        min_val = min(cpk_values)
        max_val = max(cpk_values)
        all_pass = all(v >= 1.33 for v in cpk_values)
        results["sub_items"]["3.1_CPK值"] = (
            f"✅ 通过（范围: {min_val:.2f} ~ {max_val:.2f}, 均≥1.33）"
            if all_pass
            else f"❌ 不合格（最小值 {min_val:.2f} < 1.33）"
        )
        if not all_pass:
            results["issues"].append(f"CPK值不合格：最小值 {min_val:.2f}，要求 ≥ 1.33")
    else:
        results["sub_items"]["3.1_CPK值"] = "⏱ 未检测到CPK数据"
        results["issues"].append("未在PDF中找到CPK值数据")

    # 3.2 CPK报告与图纸/全尺寸报告尺寸对应性
    has_cpk_page = any(p.get("is_cpk") for p in page_analysis)
    has_drawing = any(p.get("is_engineering_drawing") or p.get("drawing_number") for p in page_analysis)
    results["sub_items"]["3.2_CPK尺寸对应性"] = (
        "✅ 已检测到CPK及工程图纸页面（详细对应性需人工确认）"
        if has_cpk_page and has_drawing
        else "⏱ 需人工确认CPK与图纸对应关系"
    )

    has_critical_issue = any("CPK值不合格" in issue or "未找到CPK" in issue for issue in results["issues"])
    results["overall_status"] = "✅ 通过" if not has_critical_issue else "❌ 不合格"

    return results


def inspect_dimension_correspondence(page_analysis, standards):
    """
    第四类：尺寸公差对应性检验（包含关系 C ⊆ B ⊆ A）
    """
    results = {
        "sub_items": {},
        "overall_status": "⏱ 待人工确认",
        "issues": [],
    }

    has_drawing = any(p.get("is_engineering_drawing") or p.get("drawing_number") for p in page_analysis)
    has_fullsize = any(
        "全尺寸" in p.get("text", "") or "full size" in p.get("text", "").lower() or "measurement" in p.get("text", "").lower()
        for p in page_analysis
    )
    has_cpk = any(p.get("is_cpk") for p in page_analysis)

    results["sub_items"]["4.1_规格图纸(A层)"] = (
        "✅ 已提供" if has_drawing else "❌ 缺失"
    )
    results["sub_items"]["4.2_全尺寸量测报告(B层)"] = (
        "✅ 已提供" if has_fullsize else "❌ 缺失"
    )
    results["sub_items"]["4.3_CPK报告(C层)"] = (
        "✅ 已提供" if has_cpk else "❌ 缺失"
    )

    all_present = has_drawing and has_fullsize and has_cpk
    if all_present:
        results["sub_items"]["4.4_包含关系C⊆B⊆A"] = (
            "⏱ 三文件均存在，建议人工核对具体尺寸及公差是否一致"
        )
        results["overall_status"] = "⚠️ 需人工确认"
        results["issues"].append("三文件均存在，需人工核对尺寸及公差对应关系")
    else:
        missing = []
        if not has_drawing:
            missing.append("规格图纸")
        if not has_fullsize:
            missing.append("全尺寸量测报告")
        if not has_cpk:
            missing.append("CPK报告")
        results["sub_items"]["4.4_包含关系C⊆B⊆A"] = f"❌ 文件缺失，无法验证包含关系"
        results["overall_status"] = "❌ 不合格"
        results["issues"].append(f"缺失文件：{'、'.join(missing)}，无法验证尺寸包含关系")

    return results


def inspect_report_validity(page_analysis, standards, check_date):
    """
    第五类：报告时效性检验
    """
    results = {
        "sub_items": {},
        "overall_status": "✅ 通过",
        "issues": [],
    }

    all_text = " ".join(p.get("text", "") for p in page_analysis)
    all_dates = extract_dates_from_text(all_text)

    if all_dates:
        latest_date = max(all_dates)
        days_ago = (check_date.date() - latest_date).days

        if days_ago <= 365:
            results["sub_items"]["5.1_RoHS报告时效性"] = (
                f"✅ 有效（报告日期: {latest_date}，距今 {days_ago} 天 ≤ 365 天）"
            )
        else:
            results["sub_items"]["5.1_RoHS报告时效性"] = (
                f"❌ 已过期（报告日期: {latest_date}，距今 {days_ago} 天 > 365 天）"
            )
            results["issues"].append(f"RoHS报告已过期 {days_ago - 365} 天")
            results["overall_status"] = "❌ 不合格"

        # REACH报告
        has_reach = any(p.get("is_reach_survey") for p in page_analysis)
        if has_reach:
            results["sub_items"]["5.2_REACH报告时效性"] = (
                f"✅ 有效（同上）" if days_ago <= 365 else f"❌ 已过期"
            )
        else:
            results["sub_items"]["5.2_REACH报告时效性"] = "⏱ 未检测到REACH报告"
    else:
        results["sub_items"]["5.1_RoHS报告时效性"] = "⏱ 未检测到报告日期"
        results["sub_items"]["5.2_REACH报告时效性"] = "⏱ 未检测到报告日期"

    return results


# ============================================================
# V6.2 新增：目录勾选状态检测 + 料号&物料名称跨表一致性检查
# ============================================================

def extract_cover_info(page_analysis, pdf_path):
    """
    V6.2: 从封面/样品承认书页提取料号和物料名称
    返回: {"part_number": str, "material_name": str, "page_num": int}
    """
    result = {"part_number": "", "material_name": "", "page_num": 0}

    # 寻找封面页或样品承认书页
    cover_pages = [p for p in page_analysis if p.get("is_cover")]
    if not cover_pages:
        # 尝试从含物料信息的页面提取
        cover_pages = [p for p in page_analysis if any(
            kw in p.get("text", "") for kw in [
                "material number", "物料编号", "product name", "产品名称",
                "sample acknowledgement", "样品承认书", "part number", "料号"
            ]
        )]

    for p in cover_pages:
        text = p.get("text", "")
        if not text:
            continue

        lines = text.split('\n')
        for i, line in enumerate(lines):
            line_stripped = line.strip()

            # 提取料号（多种格式）
            pn_patterns = [
                r'(?:Material\s*(?:number|编号)|物料编号|Part\s*Number|料号|零件号)[\s:：]*([A-Za-z0-9_\-]+)',
                r'(?:Product\s*name|产品名称)[\s:：]*(.+)',
            ]

            # 料号匹配
            pn_match = re.search(
                r'(?:Material\s*(?:number|编号)|物料编号|Part\s*Number[（(]?料号[）)]?|料号|零件号)[\s:：\s]*([A-Za-z0-9_\-]+)',
                line, re.IGNORECASE
            )
            if pn_match and not result["part_number"]:
                result["part_number"] = pn_match.group(1).strip()
                result["page_num"] = p["page_num"]

            # 物料名称匹配
            name_match = re.search(
                r'(?:Product\s*(?:name|名称)|产品名称|Description|零件名称|物料名称|Part\s*name\s*/?\s*model|零件名称/型号)[\s:：\s]*(.+)',
                line, re.IGNORECASE
            )
            if name_match and not result["material_name"]:
                name_val = name_match.group(1).strip()
                # 过滤掉太短或不像名称的值
                if len(name_val) >= 3 and not name_val.isdigit():
                    result["material_name"] = name_val

        if result["part_number"] or result["material_name"]:
            break

    return result


def extract_table_headers_part_info(page_analysis, pdf_path):
    """
    V6.2: 从各表头提取料号和物料名称
    返回: list of dict, 每个dict代表一个表格的料号信息
    格式: [{"table_type": str, "part_number": str, "material_name": str, "page_num": int}, ...]
    """
    table_infos = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                tables = page.extract_tables()

                if not tables:
                    continue

                for table in tables:
                    if not table or len(table) < 1:
                        continue

                    # 判断表格类型（基于第一个非空行的内容）
                    table_type = "unknown"
                    first_rows_text = ""
                    for row in table[:3]:  # 检查前3行
                        if row:
                            row_str = " ".join(str(c) if c else "" for c in row)
                            first_rows_text += row_str + " "

                    first_rows_lower = first_rows_text.lower()

                    if "rohs" in first_rows_lower and ("survey" in first_rows_lower or "调查表" in first_rows_text):
                        table_type = "RoHS调查表"
                    elif "rohs" in first_rows_lower and ("test report" in first_rows_lower or "测试报告" in first_rows_text):
                        table_type = "RoHS测试报告"
                    elif "reach" in first_rows_lower:
                        table_type = "REACH报告"
                    elif "cpk" in first_rows_lower:
                        table_type = "CPK报告"
                    elif "full size" in first_rows_lower or "全尺寸" in first_rows_text:
                        table_type = "全尺寸测量报告"
                    elif "process flow" in first_rows_lower or "制造流程图" in first_rows_text:
                        table_type = "制造流程图"
                    elif "sample acknowledgement" in first_rows_lower or "样品承认书" in first_rows_text:
                        table_type = "样品承认书"

                    # 从表格中提取料号和物料名称
                    part_number = ""
                    material_name = ""

                    for row in table[:5]:  # 只查前5行（通常表头区域）
                        if not row:
                            continue
                        for cell in row:
                            if cell is None:
                                continue
                            cell_str = str(cell).strip()

                            # 匹配含料号的单元格
                            # 常见表头格式：Part Number(料号) | K5311000041LA
                            # 或：零件件号 | K5311000041LA
                            # 或跨行：Part material number 零件件号 | K5311000041LA
                            pn_cell_patterns = [
                                r'^[A-Za-z]{1,2}\d{8,}[\w\-]*$',  # 纯料号格式如 K5311000041LA
                                r'^(K|M|S|NC|XC)[A-Za-z0-9_\-]{6,}$',  # 以常见前缀开头的料号
                            ]
                            for pat in pn_cell_patterns:
                                if re.match(pat, cell_str):
                                    part_number = cell_str
                                    break

                            # 匹配含物料名称的单元格
                            # 常见格式：S1551_WIFI&BT天线_V1.1
                            name_cell_patterns = [
                                r'^[A-Za-z0-9_\-\&\+]+(?:天线|线材|屏|LCD|OLED|Display|Scanner|Speaker|Connector|Cable|模块|模组)[_Vv]?\d*\.?\d*$',
                                r'^[A-Za-z0-9_\-\&\+_Vv]+\d+\.\d+$',  # 带_Vx.y版本号的名称
                                r'^S\d+[_\w&\+\-]+$',  # S开头的产品名
                            ]
                            for pat in name_cell_patterns:
                                if re.match(pat, cell_str, re.IGNORECASE):
                                    material_name = cell_str
                                    break

                            # 备用：从键值对格式的单元格中提取
                            # 如 "Part Number(料号)" 后面紧跟值的行
                            if re.search(r'(?:part\s*number|料号|零件号|零件件号|material\s*number)', cell_str, re.IGNORECASE):
                                # 这个单元格是标签，值可能在同行下一个单元格
                                pass

                    # 如果上面的精确模式没找到，用更宽松的方式在表头区域搜索
                    if not part_number or not material_name:
                        header_text = ""
                        for row in table[:4]:
                            if row:
                                header_text += " ".join(str(c).strip() if c else "" for c in row) + " "

                        # 宽松提取料号（字母+数字组合，长度>=10）
                        if not part_number:
                            loose_pn = re.search(
                                r'(?:part\s*(?:number|no)|料号|零件号|零件件号|material\s*number)[\s:：()（）\w]*[\s:：]*([A-Za-z][A-Za-z0-9_\-]{7,})',
                                header_text, re.IGNORECASE
                            )
                            if loose_pn:
                                part_number = loose_pn.group(1)

                        # 宽松提取物料名称（中文或带下划线的英文名）
                        if not material_name:
                            loose_name = re.search(
                                r'(?:product\s*name|产品名称|description|零件名称|物料名称|part\s*name\s*[/／]\s*model|零件名称[/／]型号)[\s:：()（）\w]*[\s:：]*([A-Za-z0-9_\-&+\u4e00-\u9fff]+)',
                                header_text, re.IGNORECASE
                            )
                            if loose_name:
                                name_candidate = loose_name.group(1).strip()
                                if len(name_candidate) >= 3:
                                    material_name = name_candidate

                    if part_number or material_name:
                        table_infos.append({
                            "table_type": table_type,
                            "part_number": part_number,
                            "material_name": material_name,
                            "page_num": page.page_number,
                        })

    except Exception:
        pass

    return table_infos


def check_catalog_checkboxes(pdf_path, page_analysis):
    """
    V6.2 功能1：目录/Catalog勾选状态检测
    判定目录中的16项是否全部已勾选
    返回: dict with checklist status
    """
    result = {
        "has_catalog_page": False,
        "catalog_page_num": 0,
        "total_items": 16,
        "checked_count": 0,
        "unchecked_count": 16,
        "checked_items": [],
        "unchecked_items": [],
        "status": "⏱ 未检测到目录",
        "issues": [],
        "details": "",
    }

    # 寻找目录页
    catalog_page = None
    for p in page_analysis:
        text = p.get("text", "")
        if ("catalog" in text.lower() or "目录" in text) and \
           ("bill of material" in text.lower() or "物料清单" in text):
            catalog_page = p
            break

    if not catalog_page:
        # 尝试用关键词找目录页
        for p in page_analysis:
            text = p.get("text", "")
            if "catalog" in text.lower() or "目录" in text:
                items_found = 0
                for kw in ["bill of material", "engineering drawing", "sample photos",
                           "full size", "cpk report", "product specification",
                           "process flow", "packaging"]:
                    if kw in text.lower():
                        items_found += 1
                if items_found >= 5:  # 至少包含5个目录项才判定为目录页
                    catalog_page = p
                    break

    if not catalog_page:
        return result

    result["has_catalog_page"] = True
    result["catalog_page_num"] = catalog_page["page_num"]
    text = catalog_page["text"]

    # 方法1：通过文本特征判断勾选状态
    # PDF中的checkbox可能以 ☑(checked) / ☐(unchecked) / ✓ / □ / [x] / [ ] 等形式出现
    # 也可能是 Unicode checkbox 字符

    # 统计目录项数量（按序号1-16）
    catalog_items = []
    item_pattern = r'(?:^|\n)\s*(\d{1,2})\s+(.{10,60})'

    # 更健壮的方法：查找所有目录项文本
    expected_items = [
        "Bill of material", "Engineering drawings", "Sample photos",
        "Full size measurement report", "Cpk Report", "Product specifications",
        "Process Flow Chart", "Packaging method", "QC Flow Chart",
        "Electrical performance test report", "Reliability test report",
        "Material certificate", "RoHS 2.0 restricted substances composition questionnaire",
        "RoHS 2.0 Test Report", "REACH", "REACH Test Report"
    ]
    expected_cn_items = ["物料清单", "工程图纸", "样品照片", "全尺寸测量报告",
                         "Cpk报告", "产品规格书", "制造流程图", "包装方式",
                         "QC工程图", "电气性能测试报告", "可靠性测试报告",
                         "材质证明", "RoHS", "REACH"]

    checked_indicators = ['☑', '☒', '✓', '✔', '✅', '[x]', '[X]', '(x)', '(X)', '√']
    unchecked_indicators = ['☐', '□', '✗', '✖', '❌', '[ ]', '( )', '○']

    # 分析每一行的勾选状态
    lines = text.split('\n')
    checked = 0
    unchecked = 0
    item_details = []

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue

        # 检查是否为目录项行（含序号+项目名称）
        is_catalog_item = False
        item_name = ""

        for item in expected_items + expected_cn_items:
            if item.lower() in line_stripped.lower():
                is_catalog_item = True
                item_name = item
                break

        # 也匹配数字序号格式
        if not is_catalog_item:
            num_match = re.match(r'^(\d{1,2})\s+', line_stripped)
            if num_match:
                is_catalog_item = True
                item_name = line_stripped[:50]

        if is_catalog_item:
            # 判断此行的勾选状态
            has_checked = any(ind in line for ind in checked_indicators)
            has_unchecked = any(ind in line for ind in unchecked_indicators)

            # 特殊处理：如果同时没有明确的checked/unchecked标记，
            # 尝试从行的末尾或特定位置判断
            if not has_checked and not has_unchecked:
                # PDF中未勾选的checkbox常显示为 ☐ 或空方框
                # 已勾选的可能有 ✓ 或其他标记
                if '☐' in line or '□' in line:
                    has_unchecked = True
                elif '☑' in line:
                    has_checked = True
                else:
                    # 无法确定，默认为未勾选（保守策略）
                    has_unchecked = True

            if has_checked:
                checked += 1
                item_details.append({"item": item_name, "status": "✅ 已勾选"})
                result["checked_items"].append(item_name)
            else:
                unchecked += 1
                item_details.append({"item": item_name, "status": "❌ 未勾选"})
                result["unchecked_items"].append(item_name)

    result["checked_count"] = checked
    result["unchecked_count"] = unchecked
    result["total_items"] = checked + unchecked if (checked + unchecked) > 0 else 16

    # 判定结果
    if result["total_items"] == 0:
        result["status"] = "⏱ 未检测到目录项"
        result["details"] = "未能解析目录中的勾选框，建议人工确认"
    elif unchecked == 0 and checked > 0:
        result["status"] = "✅ 全部已勾选"
        result["details"] = f"目录共{checked}项，全部已勾选"
    elif unchecked == result["total_items"]:
        # 全部未勾选 → NG
        result["status"] = "❌ NG - 目录全部未勾选"
        result["details"] = f"目录共{result['total_items']}项，**全部未勾选**（第{result['catalog_page_num']}页）"
        result["issues"].append(f"目录页（第{result['catalog_page_num']}页）：{result['total_items']}项全部未勾选")
    elif unchecked > 0:
        # 部分未勾选 → 警告
        result["status"] = f"⚠️ 部分未勾选（{unchecked}/{result['total_items']}）"
        result["details"] = f"目录共{result['total_items']}项，已勾选{checked}项，未勾选{unchecked}项"
        result["issues"].append(f"目录页有{unchecked}项未勾选：{'、'.join(result['unchecked_items'][:5])}...")

    return result


def check_part_number_consistency(page_analysis, pdf_path):
    """
    V6.2 功能2：料号&物料名称跨表一致性检查
    提取封面和各表头的料号/物料名称，比对一致性
    返回: dict with consistency results
    """
    result = {
        "cover_info": {},
        "table_infos": [],
        "consistency_checks": [],
        "overall_status": "✅ 一致",
        "issues": [],
    }

    # Step 1: 提取封面的料号和物料名称
    cover = extract_cover_info(page_analysis, pdf_path)
    result["cover_info"] = cover

    if not cover["part_number"] and not cover["material_name"]:
        result["overall_status"] = "⏱ 无法提取封面信息"
        result["issues"].append("无法从封面/样品承认书页提取料号或物料名称")
        return result

    # Step 2: 提取各表头的料号和物料名称
    table_infos = extract_table_headers_part_info(page_analysis, pdf_path)
    result["table_infos"] = table_infos

    if not table_infos:
        result["overall_status"] = "⚠️ 未检测到其他表格的表头信息"
        result["issues"].append("未在各报告表头中找到料号/物料名称信息")
        return result

    # Step 3: 逐一比对
    ref_pn = cover["part_number"]
    ref_name = cover["material_name"]

    for ti in table_infos:
        table_type = ti["table_type"]
        table_pn = ti["part_number"]
        table_name = ti["material_name"]
        page_num = ti["page_num"]

        check_result = {
            "table_type": table_type,
            "page_num": page_num,
            "pn_match": None,
            "name_match": None,
            "pn_detail": "",
            "name_detail": "",
        }

        # 料号比对
        if ref_pn and table_pn:
            # 标准化比较（忽略大小写和空格）
            pn_normalized_ref = ref_pn.upper().replace(" ", "").replace("-", "")
            pn_normalized_table = table_pn.upper().replace(" ", "").replace("-", "")
            if pn_normalized_ref == pn_normalized_table:
                check_result["pn_match"] = "✅ 一致"
                check_result["pn_detail"] = f"{table_pn} == {ref_pn}"
            else:
                check_result["pn_match"] = "❌ 不一致"
                check_result["pn_detail"] = f"表头:{table_pn} ≠ 封面:{ref_pn}"
                result["issues"].append(f"[{table_type}(第{page_num}页)] 料号不一致：表头'{table_pn}' ≠ 封面'{ref_pn}'")
        elif ref_pn and not table_pn:
            check_result["pn_match"] = "⏱ 表头无料号"
            check_result["pn_detail"] = f"封面料号:{ref_pn}，表头未找到料号"
        else:
            check_result["pn_match"] = "⏱ 无参考"
            check_result["pn_detail"] = ""

        # 物料名称比对
        if ref_name and table_name:
            # 标准化比较（忽略大小写、空格、全角半角差异）
            name_normalized_ref = ref_name.upper().replace(" ", "").replace("－", "-").replace("_", "_")
            name_normalized_table = table_name.upper().replace(" ", "").replace("－", "-").replace("_", "_")

            # 允许微小差异（如全尺寸报告多了"线"字）
            if name_normalized_ref == name_normalized_table:
                check_result["name_match"] = "✅ 一致"
                check_result["name_detail"] = f"{table_name} == {ref_name}"
            elif name_normalized_ref in name_normalized_table or name_normalized_table in name_normalized_ref:
                check_result["name_match"] = "⚠️ 基本一致（微小差异）"
                diff = ""
                if len(name_normalized_ref) != len(name_normalized_table):
                    diff = f"(长度差{len(name_normalized_ref)} vs {len(name_normalized_table)})"
                check_result["name_detail"] = f"表头:'{table_name}' ≈ 封面:'{ref_name}' {diff}"
                result["issues"].append(f"[{table_type}(第{page_num}页)] 物料名称存在微小差异：表头'{table_name}' ≠ 封面'{ref_name}'")
            else:
                check_result["name_match"] = "❌ 不一致"
                check_result["name_detail"] = f"表头:{table_name} ≠ 封面:{ref_name}"
                result["issues"].append(f"[{table_type}(第{page_num}页)] 物料名称不一致：表头'{table_name}' ≠ 封面'{ref_name}'")
        elif ref_name and not table_name:
            check_result["name_match"] = "⏱ 表头无物料名称"
            check_result["name_detail"] = f"封面物料名称:{ref_name}，表头未找到"
        else:
            check_result["name_match"] = "⏱ 无参考"
            check_result["name_detail"] = ""

        result["consistency_checks"].append(check_result)

    # 判定整体状态
    has_mismatch = any("❌" in str(c.get("pn_match", "")) or "❌" in str(c.get("name_match", ""))
                       for c in result["consistency_checks"])
    has_minor_diff = any("⚠️" in str(c.get("pn_match", "")) or "⚠️" in str(c.get("name_match", ""))
                         for c in result["consistency_checks"])

    if has_mismatch:
        result["overall_status"] = "❌ 存在不一致"
    elif has_minor_diff:
        result["overall_status"] = "⚠️ 基本一致（有微小差异）"
    else:
        result["overall_status"] = "✅ 全部一致"

    return result


def generate_final_verdict_v62(material_type, all_results, standards):
    """
    V6.2 第六类：生成最终检验结论与处理建议（含目录勾选+料号一致性）
    """
    issues = []
    completeness = all_results.get("completeness", {})
    rohs = all_results.get("rohs", {})
    cpk = all_results.get("cpk", {})
    dimension = all_results.get("dimension", {})
    validity = all_results.get("validity", {})

    issues.extend(completeness.get("issues", []))
    issues.extend(rohs.get("issues", []))
    issues.extend(cpk.get("issues", []))
    issues.extend(dimension.get("issues", []))
    issues.extend(validity.get("issues", []))

    # V6.2: 新增目录勾选和料号一致性问题
    catalog_check = all_results.get("catalog_check", {})
    part_consistency = all_results.get("part_consistency", {})

    if catalog_check.get("issues"):
        issues.extend(catalog_check["issues"])

    if part_consistency.get("issues"):
        issues.extend(part_consistency["issues"])

    total_fail = completeness.get("fail_count", 0)
    critical_fail = (
        completeness["status"] == "❌ 不合格"
        or rohs["overall_status"] == "❌ 不合格"
        or cpk["overall_status"] == "❌ 不合格"
        or dimension["overall_status"] == "❌ 不合格"
        or validity["overall_status"] == "❌ 不合格"
        or catalog_check.get("status", "").startswith("❌")
        or part_consistency.get("overall_status", "").startswith("❌")
    )

    if critical_fail or len(issues) > 3:
        verdict = "❌ 不合格，退回重报"
        suggestion = "退回供应商，要求按XC-R-0802-DQM-002正式样品承认书模板重新提交完整的样品承认书\n\n**必须补充/修正以下问题：**\n"
        for item in completeness.get("items", []):
            if "❌" in item["结果"]:
                suggestion += f"- ❌ **{item['项目']}**\n"

        # V6.2: 目录勾选NG
        if catalog_check.get("status", "").startswith("❌"):
            suggestion += f"- ❌ **目录全部未勾选**（第{catalog_check['catalog_page_num']}页）：请在样品承认书目录中勾选所有已提供的文件项\n"
        elif catalog_check.get("unchecked_count", 0) > 0 and catalog_check.get("unchecked_count") < catalog_check.get("total_items", 16):
            suggestion += f"- ⚠️ **目录部分未勾选**（{catalog_check['unchecked_count']}/{catalog_check['total_items']}项未勾选）\n"

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


# 保持向后兼容
def generate_final_verdict(material_type, all_results, standards):
    """向后兼容的包装"""
    return generate_final_verdict_v62(material_type, all_results, standards)


# ============================================================
# 主审核流程（V4.0 完整8步）
# ============================================================

def run_full_inspection(file_path, file_name, standards):
    """执行完整的V4.0 6大类审核流程（含PDF逐页分析）"""
    check_date = datetime.now()
    version = standards.get("version", "未知")

    # 提取PDF并逐页分析（V4.0 核心）
    with st.spinner(f"正在逐页分析PDF: {file_name}..."):
        page_analysis = analyze_pdf_page_by_page(file_path)
        all_text = extract_all_text(file_path)

    # 第零步：物料类型判定
    mat_type, mat_type_cn, mat_kw, needs_elec = determine_material_type(file_name, standards)

    # 第一步：文件类型检查（V4.0新增）
    file_type, file_type_note = check_file_type(file_path, all_text)

    # 第二步~第五步：各类检验（基于逐页分析结果）
    completeness = inspect_file_completeness_v4(page_analysis, mat_type, standards)
    rohs = inspect_rohs_compliance(page_analysis, standards, check_date)
    cpk = inspect_cpk_compliance(page_analysis, standards, all_text, pdf_path=file_path)
    dimension = inspect_dimension_correspondence(page_analysis, standards)
    validity = inspect_report_validity(page_analysis, standards, check_date)

    # V6.2 新增：目录勾选状态检测 + 料号跨表一致性检查
    catalog_check = check_catalog_checkboxes(file_path, page_analysis)
    part_consistency = check_part_number_consistency(page_analysis, file_path)

    # 第六步：生成最终结论
    all_results = {
        "completeness": completeness,
        "rohs": rohs,
        "cpk": cpk,
        "dimension": dimension,
        "validity": validity,
        "catalog_check": catalog_check,
        "part_consistency": part_consistency,
    }
    final = generate_final_verdict_v62(mat_type, all_results, standards)

    # 构建返回结果
    result = {
        "文件名": file_name,
        "文件类型": file_type_note,
        "物料类型": f"{mat_type_cn}" if mat_type != "unknown" else "未知（需人工确认）",
        "需要电气性能测试": "是" if needs_elec else "否",
        "文件完整性": completeness["status"],
        "RoHS合规性": rohs["overall_status"],
        "CPK合规性": cpk["overall_status"],
        "尺寸对应性": dimension["overall_status"],
        "报告时效性": validity["overall_status"],
        # V6.2 新增
        "目录勾选状态": catalog_check.get("status", "⏱ 未检测"),
        "料号一致性": part_consistency.get("overall_status", "⏱ 未检测"),
        # 原有
        "总体结论": final["verdict"],
        "问题数量": final["issue_count"],
        "审核时间": check_date.strftime("%Y-%m-%d %H:%M:%S"),
        "标准版本": f"V{version}",
        "_detail": {
            "mat_type": mat_type,
            "mat_type_cn": mat_type_cn,
            "needs_elec": needs_elec,
            "file_type": file_type,
            "file_type_note": file_type_note,
            "page_analysis": [
                {
                    "page": p["page_num"],
                    "is_drawing": p.get("is_engineering_drawing"),
                    "drawing_number": p.get("drawing_number"),
                    "is_bom": p.get("is_bom"),
                    "is_cpk": p.get("is_cpk"),
                    "is_rohs": p.get("is_rohs_survey"),
                    "is_reach": p.get("is_reach_survey"),
                }
                for p in page_analysis if "error" not in p
            ],
            "completeness": completeness,
            "rohs": rohs,
            "cpk": cpk,
            "dimension": dimension,
            "validity": validity,
            # V6.2 新增
            "catalog_check": catalog_check,
            "part_consistency": part_consistency,
            #
            "final": final,
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

# 从标准文件动态生成选项
electronic_items = standards.get("file_completeness", {}).get("electronic", {}).get("items", [])
all_item_names = [item["name"] for item in electronic_items]
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
# V6.2 新增
catalog_check = st.sidebar.checkbox("✅ 目录勾选状态检测（V6.2）", value=True)
part_check = st.sidebar.checkbox("✅ 料号&物料名称跨表一致性（V6.2）", value=True)

min_cpk_display = 1.33
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
        all_paths = []
        if uploaded_files:
            for uf in uploaded_files:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                tmp.write(uf.getvalue())
                all_paths.append((tmp.name, uf.name))
        if 'folder_files' in st.session_state:
            for fp in st.session_state['folder_files']:
                all_paths.append((fp, os.path.basename(fp)))

        if not all_paths:
            st.warning("⚠️ 请先上传或选择PDF文件")
        else:
            st.info(f"开始审核 **{len(all_paths)}** 个文件...")

            progress = st.progress(0)
            detail_results = []

            for i, (fp, fname) in enumerate(all_paths):
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
                    st.markdown(f"**文件类型:** {d['file_type_note']}")
                    st.markdown(f"**物料类型:** {d['mat_type_cn']} | **需电气性能测试:** {'是' if d['needs_elec'] else '否'}")

                    # PDF页面结构分析（V4.0新增）
                    st.subheader("📄 PDF页面结构分析（V4.0）")
                    if d.get("page_analysis"):
                        page_df = pd.DataFrame(d["page_analysis"])
                        st.dataframe(page_df, use_container_width=True, hide_index=True)
                    else:
                        st.text("⏱️ 无页面结构分析数据")

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

                    # V6.2 新增：目录勾选状态
                    if d.get("catalog_check"):
                        cat = d["catalog_check"]
                        st.subheader("6️⃣ 目录/Catalog勾选状态（V6.2）")
                        st.markdown(f"**判定结果:** {cat.get('status', 'N/A')}")
                        if cat.get('details'):
                            st.text(cat['details'])
                        if cat.get("has_catalog_page"):
                            col_cat1, col_cat2 = st.columns(2)
                            with col_cat1:
                                st.metric("已勾选", cat.get("checked_count", 0))
                            with col_cat2:
                                st.metric("未勾选", cat.get("unchecked_count", 0))
                        if cat.get("unchecked_items"):
                            st.text(f"未勾选项: {', '.join(cat['unchecked_items'][:10])}")

                    # V6.2 新增：料号&物料名称跨表一致性
                    if d.get("part_consistency"):
                        pc = d["part_consistency"]
                        st.subheader("7️⃣ 料号&物料名称跨表一致性（V6.2）")
                        st.markdown(f"**整体判定:** {pc.get('overall_status', 'N/A')}")
                        # 显示封面参考信息
                        cover = pc.get("cover_info", {})
                        if cover.get("part_number") or cover.get("material_name"):
                            st.markdown(f"**封面参考信息:** 料号=`{cover.get('part_number', '未提取')}` | 名称=`{cover.get('material_name', '未提取')}`")
                        # 一致性检查表格
                        consistency_checks = pc.get("consistency_checks", [])
                        if consistency_checks:
                            pc_rows = []
                            for cc in consistency_checks:
                                pc_rows.append({
                                    "报告类型": cc["table_type"],
                                    "页码": cc["page_num"],
                                    "料号匹配": cc.get("pn_match", ""),
                                    "名称匹配": cc.get("name_match", ""),
                                })
                            pc_df = pd.DataFrame(pc_rows)
                            st.dataframe(pc_df, use_container_width=True, hide_index=True)

                    # 最终处理建议
                    st.subheader("8️⃣ 检验结论与处理建议")
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
                    full_report += f"**文件类型:** {d['file_type_note']}\n"
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
            for fp, _ in all_paths:
                if fp.startswith(tempfile.gettempdir()):
                    try:
                        os.unlink(fp)
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

### ⚙️ V6.2 版本说明（与SKILL.md同步）
| 功能 | 状态 |
|------|--------|
| 文件类型检查（DQM-001 vs DQM-002） | ✅ |
| PDF逐页结构分析 | ✅ |
| 工程图纸判定规则（图号/版本/尺寸标注/标题栏） | ✅ |
| 产品规格书判定规则（嵌入工程图纸的情况） | ✅ |
| 电子料/结构件区分审核 | ✅ |
| RoHS合规（4子项，含红框6字段） | ✅ |
| CPK合规（2子项，≥1.33，支持表格提取） | ✅ |
| 尺寸对应性（包含关系C⊆B⊆A） | ✅ |
| 报告时效性（≤1年） | ✅ |
| **目录勾选状态检测** | ✅ **V6.2新增** |
| **料号&物料名称跨表一致性检查** | ✅ **V6.2新增** |

> 💡 **注意：** 本工具使用PDF文本解析技术进行自动化审核，部分内容（如图片中的尺寸标注、手写签名等）可能需要人工辅助确认。
""")

hide_style = "<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;}</style>"
st.markdown(hide_style, unsafe_allow_html=True)
