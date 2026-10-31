import os
import re
import sys
import subprocess
from copy import copy
from io import BytesIO
from datetime import datetime

# ==========================================
# 0. 自动引导启动器
# ==========================================
if __name__ == "__main__":
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        in_streamlit = get_script_run_ctx() is not None
    except ImportError:
        in_streamlit = False

    if not in_streamlit:
        script_path = os.path.abspath(__file__)
        subprocess.run([sys.executable, "-m", "streamlit", "run", script_path])
        sys.exit(0)

import streamlit as st
import pandas as pd
import openpyxl


# ==========================================
# 1. 资源路径兼容函数
# ==========================================
def get_resource_path(relative_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, relative_path)


TEMPLATE_PATH = get_resource_path("Quotation Template.xlsx")
LOGO_PATH = get_resource_path("Tofflon Logo.png")


# ==========================================
# 2. 产品部报价单智能解析引擎
# ==========================================
def detect_and_parse_source_excel(uploaded_file) -> list:
    file_bytes = uploaded_file.getvalue()
    wb_val = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    sheet_names = wb_val.sheetnames
    
    if len(sheet_names) > 1 and ("List" in sheet_names or any("线" in s or "DP" in s for s in sheet_names)):
        target_sheet_name = next((s for s in sheet_names if s != "List"), sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])
        ws_val = wb_val[target_sheet_name]
    else:
        ws_val = wb_val.active

    header_row = None
    no_col = None
    for r in range(1, min(25, ws_val.max_row + 1)):
        for c in range(1, min(10, ws_val.max_column + 1)):
            cell_val = str(ws_val.cell(r, c).value or "").strip().upper()
            if cell_val == "NO.":
                header_row = r
                no_col = c
                break
        if header_row:
            break

    if not header_row:
        raise ValueError("未能识别该表格的表头（未找到 'NO.' 列）。")

    rmb_price_col = None
    for c in range(no_col, ws_val.max_column + 1):
        h_val = str(ws_val.cell(header_row, c).value or "").strip().upper()
        if "STANDARD PRICE(RMB)" in h_val or "人民币" in h_val or "RMB" in h_val:
            if "SUB" not in h_val:
                rmb_price_col = c
                break

    machines = []
    current_machine = None
    current_section = None

    for r in range(header_row + 1, ws_val.max_row + 1):
        raw_no = ws_val.cell(r, no_col).value
        no_str = str(raw_no).strip() if raw_no is not None else ""
        desc_val = ws_val.cell(r, no_col + 1).value
        config_val = ws_val.cell(r, no_col + 2).value
        qty_val = ws_val.cell(r, no_col + 3).value
        opt_val = ws_val.cell(r, no_col + 4).value
        remark_val = ws_val.cell(r, no_col + 5).value
        
        rmb_val = 0
        if rmb_price_col:
            val = ws_val.cell(r, rmb_price_col).value
            if isinstance(val, (int, float)):
                rmb_val = val
        if rmb_val == 0:
            g_val = ws_val.cell(r, no_col + 6).value
            if isinstance(g_val, (int, float)):
                rmb_val = g_val

        no_upper = no_str.upper()

        if any(no_upper.startswith(p) for p in ["B-DOCUMENTATION", "C-SERVICE", "W-SERVICE", "D-DOCUMENTATION", "S-SERVICE"]):
            if current_machine:
                machines.append(current_machine)
                current_machine = None
            break

        if any(kw in no_upper for kw in ["SUB TOTAL OF THE EQUIPMENT", "DOCUMENTATION", "RELATED SERVICE", "COMMERCIAL TERMS", "A-TOTAL EQUIPMENT"]):
            if current_machine:
                machines.append(current_machine)
                current_machine = None
            break

        is_machine_title = False
        if no_str and (no_str.startswith("A ") or no_str.startswith("B-") or no_str.startswith("C-") or no_str.startswith("D-") or no_str.startswith("E-") or no_str.startswith("F-")):
            if not re.match(r"^[A-F]-\d+\.", no_str) and not re.match(r"^[A-F]-\d+$", no_str):
                is_machine_title = True

        if is_machine_title:
            if current_machine:
                machines.append(current_machine)
            current_machine = {
                "machine_no": no_str,
                "machine_desc": desc_val or no_str,
                "sections": []
            }
            current_section = None
            continue

        if not current_machine:
            current_machine = {
                "machine_no": "A-Machine A",
                "machine_desc": "Equipment",
                "sections": []
            }

        if re.match(r"^[A-Z]-\d+$", no_str):
            current_section = {
                "section_no": no_str,
                "section_title": desc_val or "Main Machine / Options",
                "items": []
            }
            current_machine["sections"].append(current_section)
            continue

        if no_str and re.match(r"^[A-Z]-\d+\.\d+", no_str):
            if not current_section:
                current_section = {
                    "section_no": f"{no_str.split('-')[0]}-1",
                    "section_title": "Main Machine",
                    "items": []
                }
                current_machine["sections"].append(current_section)

            current_section["items"].append({
                "no": no_str,
                "model": desc_val,
                "config": config_val,
                "qty": qty_val if isinstance(qty_val, (int, float)) else 1,
                "options": opt_val,
                "remark": remark_val,
                "rmb_base_price": rmb_val,
            })

    if current_machine and current_machine not in machines:
        machines.append(current_machine)

    return machines if machines else [{"machine_no": "A-Machine A", "machine_desc": "Equipment", "sections": []}]


# ==========================================
# 3. OpenPyXL 样式与动态构建引擎
# ==========================================
def capture_row_style(ws, row_idx: int, max_col: int = 10) -> dict:
    cell_styles = {}
    for c in range(1, max_col + 1):
        cell = ws.cell(row_idx, c)
        cell_styles[c] = {
            "font": copy(cell.font), "fill": copy(cell.fill),
            "border": copy(cell.border), "alignment": copy(cell.alignment),
            "number_format": cell.number_format,
        }
    return {"height": ws.row_dimensions[row_idx].height, "cells": cell_styles}


def apply_row_style(ws, target_row: int, style_prototype: dict):
    ws.row_dimensions[target_row].height = style_prototype["height"]
    for c, st_dict in style_prototype["cells"].items():
        cell = ws.cell(target_row, c)
        cell.font = copy(st_dict["font"])
        cell.fill = copy(st_dict["fill"])
        cell.border = copy(st_dict["border"])
        cell.alignment = copy(st_dict["alignment"])
        cell.number_format = st_dict["number_format"]


def snapshot_tail_rows(ws, start_row: int, end_row: int, max_col: int = 10) -> dict:
    rows_data = []
    for r in range(start_row, end_row + 1):
        row_cells = {c: ws.cell(r, c).value for c in range(1, max_col + 1)}
        rows_data.append({"orig_row": r, "values": row_cells, "style": capture_row_style(ws, r, max_col)})

    merges = [{"min_row": m.min_row, "max_row": m.max_row, "min_col": m.min_col, "max_col": m.max_col} 
              for m in ws.merged_cells.ranges if m.min_row >= start_row]
    return {"start_row": start_row, "rows": rows_data, "merges": merges}


def shift_excel_formula(formula: str, row_delta: int) -> str:
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    def _repl(match):
        col_str, row_num = match.group(1), int(match.group(2))
        return f"{col_str}{row_num + row_delta}" if row_num >= 56 else match.group(0)
    return re.sub(r"([A-Z]+)(\d+)", _repl, formula)


def build_quotation_excel(header_info: dict, machines_data: list | None, custom_docs: list | None, service_data: dict | None, custom_services: list | None) -> BytesIO:
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["Template"]

    # 1. 写入表头 Rows 8 - 12
    ws["B8"].value = f"Customer: {header_info['customer']}"
    ws["B9"].value = f"Customer Contacts: {header_info['customer_contacts']}"
    ws["G9"].value = f"Quotation No.: {header_info['quotation_no']}"
    ws["B10"].value = f"Tofflon Contacts: {header_info['tofflon_contacts']}"
    ws["G10"].value = f"Date: {header_info['today_str']}"
    ws["B11"].value = f"Incoterms: {header_info['incoterms_str']}"
    ws["B12"].value = "Currency:"
    ws["J12"].value = header_info["currency"]

    if not machines_data:
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    style_level1 = capture_row_style(ws, 14)
    style_level2 = capture_row_style(ws, 15)
    style_item = capture_row_style(ws, 16)
    style_sub_sub = capture_row_style(ws, 29)
    style_sub_main = capture_row_style(ws, 33)

    orig_tail_start = 56
    orig_max_row = ws.max_row
    tail_snapshot = snapshot_tail_rows(ws, orig_tail_start, orig_max_row)

    merges_to_remove = [m for m in list(ws.merged_cells.ranges) if m.min_row >= 14]
    for m in merges_to_remove:
        ws.unmerge_cells(str(m))

    for r in range(14, orig_max_row + 200):
        for c in range(1, 12):
            ws.cell(r, c).value = None

    curr_row = 14
    machine_subtotal_rows = []
    machine_letters = ["A", "B", "C", "D", "E", "F", "G", "H"]
    
    markup_factor = 1.0 - (header_info["markup_pct"] / 100.0)
    exchange_rate = header_info["exchange_rate"]

    # 2. 动态生成 Machine 区块
    for idx, mach in enumerate(machines_data):
        m_letter = machine_letters[idx] if idx < len(machine_letters) else "A"
        
        apply_row_style(ws, curr_row, style_level1)
        ws.cell(curr_row, 2).value = f"{m_letter}-Machine {m_letter}"
        ws.cell(curr_row, 4).value = mach.get("machine_desc") or f"Equipment System {m_letter}"
        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=3)
        ws.merge_cells(start_row=curr_row, start_column=4, end_row=curr_row, end_column=10)
        curr_row += 1

        sub_rows_in_mach = []

        for s_idx, sec in enumerate(mach["sections"]):
            apply_row_style(ws, curr_row, style_level2)
            ws.cell(curr_row, 2).value = f"{m_letter}-{s_idx+1}"
            ws.cell(curr_row, 3).value = sec["section_title"]
            curr_row += 1

            sec_start = curr_row
            for item in sec["items"]:
                apply_row_style(ws, curr_row, style_item)
                ws.cell(curr_row, 2).value = item["no"]
                ws.cell(curr_row, 3).value = item["model"]
                ws.cell(curr_row, 4).value = item["config"]
                ws.cell(curr_row, 5).value = item["qty"]
                ws.cell(curr_row, 6).value = item["options"]
                ws.cell(curr_row, 7).value = item["remark"]
                
                rmb_base = item["rmb_base_price"]
                if rmb_base > 0 and exchange_rate > 0 and markup_factor > 0:
                    unit_price_formula = f"=ROUNDUP({rmb_base}/({exchange_rate}*{markup_factor}),-2)"
                    ws.cell(curr_row, 8).value = unit_price_formula
                else:
                    ws.cell(curr_row, 8).value = 0
                
                opt_str = str(item["options"]).strip().lower()
                if "option" in opt_str:
                    ws.cell(curr_row, 9).value = f"=E{curr_row}*H{curr_row}"
                    ws.cell(curr_row, 10).value = None
                else:
                    ws.cell(curr_row, 9).value = None
                    ws.cell(curr_row, 10).value = f"=E{curr_row}*H{curr_row}"

                ws.cell(curr_row, 8).number_format = "#,##0"
                ws.cell(curr_row, 9).number_format = "#,##0"
                ws.cell(curr_row, 10).number_format = "#,##0"
                curr_row += 1
            sec_end = curr_row - 1

            apply_row_style(ws, curr_row, style_sub_sub)
            ws.cell(curr_row, 2).value = f"Sub Total of {m_letter}-{s_idx+1}"
            ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=8)
            if sec_end >= sec_start:
                ws.cell(curr_row, 9).value = f"=SUM(I{sec_start}:I{sec_end})"
                ws.cell(curr_row, 10).value = f"=SUM(J{sec_start}:J{sec_end})"
            else:
                ws.cell(curr_row, 9).value = 0
                ws.cell(curr_row, 10).value = 0
            ws.cell(curr_row, 9).number_format = "#,##0"
            ws.cell(curr_row, 10).number_format = "#,##0"
            sub_rows_in_mach.append(curr_row)
            curr_row += 1

        apply_row_style(ws, curr_row, style_sub_main)
        ws.cell(curr_row, 2).value = f"Sub Total for {m_letter}"
        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=8)
        if sub_rows_in_mach:
            ws.cell(curr_row, 9).value = f"={'+'.join([f'I{r}' for r in sub_rows_in_mach])}"
            ws.cell(curr_row, 10).value = f"={'+'.join([f'J{r}' for r in sub_rows_in_mach])}"
        else:
            ws.cell(curr_row, 9).value = 0
            ws.cell(curr_row, 10).value = 0
        ws.cell(curr_row, 9).number_format = "#,##0"
        ws.cell(curr_row, 10).number_format = "#,##0"
        machine_subtotal_rows.append(curr_row)
        curr_row += 1

    # 3. 动态拼接尾部区块 (W区、V区及商务条款)
    w_start_row = curr_row
    row_delta = curr_row - orig_tail_start
    custom_docs_count = len(custom_docs) if custom_docs else 0
    custom_services_count = len(custom_services) if custom_services else 0
    
    for r_info in tail_snapshot["rows"]:
        orig_r = r_info["orig_row"]
        new_r = orig_r + row_delta
        
        apply_row_style(ws, new_r, r_info["style"])
        for c, val in r_info["values"].items():
            cell_val = shift_excel_formula(val, row_delta)
            
            # 自动替换 V 区配置文本中的 [x] 和 [y]
            if service_data and isinstance(cell_val, str) and ("[x]" in cell_val or "[y]" in cell_val):
                v_map = {
                    67: service_data["items"]["fat"],
                    68: service_data["items"]["unpacking"],
                    69: service_data["items"]["positioning"],
                    70: service_data["items"]["restore"],
                    71: service_data["items"]["commissioning"],
                    72: service_data["items"]["sat"],
                    73: service_data["items"]["iq_oq"],
                    74: service_data["items"]["training"],
                }
                if orig_r in v_map:
                    info = v_map[orig_r]
                    cell_val = cell_val.replace("[x]", str(info["engineer"])).replace("[y]", str(info["days"]))

            # 自动替换商务条款 Delivery 中的 [x] 交期
            if isinstance(cell_val, str) and "[x]" in cell_val and orig_r == 80:
                cell_val = cell_val.replace("[x]", str(header_info["delivery_months"]))

            ws.cell(new_r, c).value = cell_val

        # W-2 与 W-3 手动定价写入
        if orig_r == 58:
            ws.cell(new_r, 10).value = header_info["w2_price"]
            ws.cell(new_r, 10).number_format = "#,##0"
        elif orig_r == 59:
            ws.cell(new_r, 10).value = header_info["w3_price"]
            ws.cell(new_r, 10).number_format = "#,##0"

    w_end_row = 64 + row_delta
    
    # 4. 动态追加 W 区自定义文件行
    if custom_docs_count > 0:
        ws.insert_rows(w_end_row + 1, amount=custom_docs_count)
        style_sample_row = capture_row_style(ws, w_end_row)
        for idx, doc in enumerate(custom_docs):
            insert_r = w_end_row + 1 + idx
            apply_row_style(ws, insert_r, style_sample_row)
            ws.cell(insert_r, 2).value = f"W-{8 + idx + 1}"
            ws.cell(insert_r, 3).value = doc.get("Model/Description", "")
            ws.cell(insert_r, 4).value = doc.get("Technical Configuration", "")
            ws.cell(insert_r, 5).value = doc.get("QTY", 1)
            is_opt = str(doc.get("Option", "No")).strip().capitalize() == "Yes"
            ws.cell(insert_r, 6).value = "Options" if is_opt else None
            ws.cell(insert_r, 7).value = doc.get("Remark", "")
            unit_p = doc.get("Unit Price", 0)
            ws.cell(insert_r, 8).value = unit_p
            ws.cell(insert_r, 8).number_format = "#,##0"
            if is_opt:
                ws.cell(insert_r, 9).value = f"=E{insert_r}*H{insert_r}"
                ws.cell(insert_r, 9).number_format = "#,##0"
            else:
                ws.cell(insert_r, 10).value = f"=E{insert_r}*H{insert_r}"
                ws.cell(insert_r, 10).number_format = "#,##0"
        w_end_row += custom_docs_count

    w_sub_row = 65 + row_delta + custom_docs_count
    ws.cell(w_sub_row, 9).value = f"=SUM(I{w_start_row}:I{w_end_row})"
    ws.cell(w_sub_row, 10).value = f"=SUM(J{w_start_row}:J{w_end_row})"

    # 5. V区（Service）标准服务数据计算与填入
    v_start_row = 67 + row_delta + custom_docs_count
    v_orig_rows = {
        "fat": v_start_row,
        "unpacking": v_start_row + 1,
        "positioning": v_start_row + 2,
        "restore": v_start_row + 3,
        "commissioning": v_start_row + 4,
        "sat": v_start_row + 5,
        "iq_oq": v_start_row + 6,
        "training": v_start_row + 7,
    }

    # 显式确保 V-3 和 V-4 的 C 列名称正确填入
    ws.cell(v_start_row + 5, 3).value = "Site Acceptance Test (SAT)"
    ws.cell(v_start_row + 6, 3).value = "IQ & OQ Execution"

    fat_rate = service_data.get("fat_rate", 0)
    onsite_rate = service_data.get("onsite_rate", 0)

    for k, info in service_data["items"].items():
        if k in v_orig_rows:
            target_r = v_orig_rows[k]
            eng = info["engineer"]
            days = info["days"]
            rate = fat_rate if k == "fat" else onsite_rate
            calc_unit_price = rate * eng * days
            ws.cell(target_r, 5).value = 1
            ws.cell(target_r, 8).value = calc_unit_price
            ws.cell(target_r, 8).number_format = "#,##0"
            ws.cell(target_r, 10).value = f"=E{target_r}*H{target_r}"
            ws.cell(target_r, 10).number_format = "#,##0"

    # 6. V区动态追加自定义服务项（插在 Packing 之前）
    packing_orig_row = v_start_row + 8
    if custom_services_count > 0:
        ws.insert_rows(packing_orig_row, amount=custom_services_count)
        style_sample_v = capture_row_style(ws, packing_orig_row + custom_services_count)
        for idx, srv in enumerate(custom_services):
            insert_r = packing_orig_row + idx
            apply_row_style(ws, insert_r, style_sample_v)
            ws.cell(insert_r, 2).value = f"V-6.{idx+1}"
            ws.cell(insert_r, 3).value = srv.get("Model/Description", "")
            ws.cell(insert_r, 4).value = srv.get("Technical Configuration", "")
            ws.cell(insert_r, 5).value = srv.get("QTY", 1)
            is_opt = str(srv.get("Option", "No")).strip().capitalize() == "Yes"
            ws.cell(insert_r, 6).value = "Options" if is_opt else None
            ws.cell(insert_r, 7).value = srv.get("Remark", "")
            unit_p = srv.get("Unit Price", 0)
            ws.cell(insert_r, 8).value = unit_p
            ws.cell(insert_r, 8).number_format = "#,##0"
            if is_opt:
                ws.cell(insert_r, 9).value = f"=E{insert_r}*H{insert_r}"
                ws.cell(insert_r, 9).number_format = "#,##0"
            else:
                ws.cell(insert_r, 10).value = f"=E{insert_r}*H{insert_r}"
                ws.cell(insert_r, 10).number_format = "#,##0"
        packing_orig_row += custom_services_count

    # 7. Packing 行更新（名称固定为 Packing，描述放在 G 列，价格与 QTY=1）
    packing_row = packing_orig_row
    ws.cell(packing_row, 3).value = "Packing"  # Col C 必须保持 Packing
    ws.cell(packing_row, 5).value = 1  # QTY = 1
    ws.cell(packing_row, 7).value = header_info["packing_desc"]  # Col G: Remark 放置修改后的描述
    ws.cell(packing_row, 8).value = header_info["packing_price"]  # Unit Price
    ws.cell(packing_row, 8).number_format = "#,##0"
    ws.cell(packing_row, 10).value = f"=E{packing_row}*H{packing_row}"  # Sub Price
    ws.cell(packing_row, 10).number_format = "#,##0"

    v_end_row = packing_row
    v_sub_row = v_end_row + 1
    
    # 动态写入 V 区小计公式
    ws.cell(v_sub_row, 9).value = f"=SUM(I{v_start_row}:I{v_end_row})"
    ws.cell(v_sub_row, 10).value = f"=SUM(J{v_start_row}:J{v_end_row})"
    ws.cell(v_sub_row, 9).number_format = "#,##0"
    ws.cell(v_sub_row, 10).number_format = "#,##0"

    # 8. 修正 Total Amount 行
    for r in range(ws.min_row, ws.max_row + 1):
        if str(ws.cell(r, 2).value or "").strip() == "Total Amount":
            i_sum_parts = [f"I{r_sub}" for r_sub in machine_subtotal_rows]
            j_sum_parts = [f"J{r_sub}" for r_sub in machine_subtotal_rows]
            i_sum_parts.extend([f"I{w_sub_row}", f"I{v_sub_row}"])
            j_sum_parts.extend([f"J{w_sub_row}", f"J{v_sub_row}"])
            ws.cell(r, 9).value = f"={'+'.join(i_sum_parts)}"
            ws.cell(r, 10).value = f"={'+'.join(j_sum_parts)}"
            ws.cell(r, 9).number_format = "#,##0"
            ws.cell(r, 10).number_format = "#,##0"

    # 9. 处理尾部合并单元格平移（精准保护 W 区和 V 区的合并单元格）
    for m in tail_snapshot["merges"]:
        orig_min = m["min_row"]
        orig_max = m["max_row"]
        
        if (orig_min == 57 and orig_max == 59) or (orig_min == 68 and orig_max == 71):
            continue

        shift_offset = row_delta
        if orig_min >= 57:
            shift_offset += custom_docs_count
        if orig_min >= 67:
            shift_offset += custom_services_count

        ws.merge_cells(
            start_row=orig_min + shift_offset,
            start_column=m["min_col"],
            end_row=orig_max + shift_offset,
            end_column=m["max_col"]
        )

    # 动态灵活合并 W-1 到 W-3 的 C 列
    ws.merge_cells(start_row=w_start_row, start_column=3, end_row=w_start_row + 2, end_column=3)
    
    # 动态灵活合并 V-2.1 到 V-2.4 的 C 列
    ws.merge_cells(start_row=v_start_row + 1, start_column=3, end_row=v_start_row + 4, end_column=3)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==========================================
# 4. Streamlit 界面 (中英文双语版 + Logo)
# ==========================================
st.set_page_config(page_title="东富龙自动化报价格式转化系统 | Tofflon Quotation Formalisation System", page_icon="📄", layout="wide")

# 标题与 Logo 布局
col_title, col_logo = st.columns([4, 1])
with col_title:
    st.title("📄 东富龙自动化报价格式转化系统\nTofflon Automatic Quotation Formalisation System")
    st.caption("支持中英文双语、单选 Incoterms、交期替换、W/V区动态合并与矩阵自动计算 | Supports Bilingual, Incoterms, Delivery Substitution & Matrix Calculation")

with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=180)
    else:
        st.markdown("### **Tofflon**")

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"⚠️ 未找到标准模板文件 / Template file not found: `{TEMPLATE_PATH}`")
    st.stop()

# --- 区块 1：表头基础信息 ---
st.subheader("1. 表头基础信息 / Header Information (Rows 8 to 12)")
col_l, col_r = st.columns(2)
with col_l:
    customer = st.text_input("客户名称 / Customer", value="PolyPeptide Group")
    customer_contacts = st.text_input("客户联系人 / Customer Contacts", placeholder="联系人 / Contact Person")
    tofflon_contacts = st.text_input("东富龙联系人 / Tofflon Contacts", value="Adler Yu")
    markup_pct = st.number_input("销售价格上涨比例 / Price Markup Percentage (xx %)", min_value=0.0, max_value=99.9, value=15.0, step=1.0)

with col_r:
    quotation_no = st.text_input("报价单号 / Quotation No.", value="TFL-2026-015")
    today_str = datetime.now().strftime("%Y.%m.%d")
    st.text_input("日期 / Date (系统自动 / Auto)", value=today_str, disabled=True)
    currency = st.text_input("币种 / Currency", value="EUR")
    exchange_rate = st.number_input("手动汇率设定 / Exchange Rate (Manual)", min_value=0.0001, max_value=1000.0, value=7.0, step=0.01)

# --- 区块 2：Incoterms 贸易术语设定（单选） ---
st.subheader("2. 贸易术语设定 / Incoterms Setting (Row 11)")
INCOTERM_OPTIONS = [
    "EXW, Tofflon Plant, Shanghai, China (Incoterms 2020)",
    "FCA, Tofflon Plant, Shanghai, China (Incoterms 2020)",
    "CIP, Port, City, Country (Incoterms 2020)",
    "DAP, Customer Plant, City, Country (Incoterms 2020)",
    "DDP, Customer Plant, City, Country (Incoterms 2020)",
]

selected_incoterm_option = st.selectbox("请选择贸易术语（单选） / Select Incoterms (Single-select):", options=INCOTERM_OPTIONS, index=1)

final_incoterms_str = ""
if selected_incoterm_option.startswith("EXW"):
    final_incoterms_str = "EXW, Tofflon Plant, Shanghai, China (Incoterms 2020)"
elif selected_incoterm_option.startswith("FCA"):
    final_incoterms_str = "FCA, Tofflon Plant, Shanghai, China (Incoterms 2020)"
elif selected_incoterm_option.startswith("CIP"):
    with st.container(border=True):
        st.markdown("📍 **CIP 目的地信息填写 / CIP Destination Details**")
        c1, c2, c3 = st.columns(3)
        cip_port = c1.text_input("港口/机场 / Port", key="cip_port")
        cip_city = c2.text_input("城市 / City", key="cip_city")
        cip_country = c3.text_input("国家 / Country", key="cip_country")
        final_incoterms_str = f"CIP, {cip_port.strip() or 'Port'}, {cip_city.strip() or 'City'}, {cip_country.strip() or 'Country'} (Incoterms 2020)"
elif selected_incoterm_option.startswith("DAP"):
    with st.container(border=True):
        st.markdown("🏭 **DAP 客户工厂信息填写 / DAP Plant Details**")
        c1, c2, c3 = st.columns(3)
        dap_plant = c1.text_input("客户工厂 / Customer Plant", key="dap_plant")
        dap_city = c2.text_input("城市 / City", key="dap_city")
        dap_country = c3.text_input("国家 / Country", key="dap_country")
        final_incoterms_str = f"DAP, {dap_plant.strip() or 'Customer Plant'}, {dap_city.strip() or 'City'}, {dap_country.strip() or 'Country'} (Incoterms 2020)"
elif selected_incoterm_option.startswith("DDP"):
    with st.container(border=True):
        st.markdown("🏭 **DDP 客户工厂信息填写 / DDP Plant Details**")
        c1, c2, c3 = st.columns(3)
        ddp_plant = c1.text_input("客户工厂 / Customer Plant", key="ddp_plant")
        ddp_city = c2.text_input("城市 / City", key="ddp_city")
        ddp_country = c3.text_input("国家 / Country", key="ddp_country")
        final_incoterms_str = f"DDP, {ddp_plant.strip() or 'Customer Plant'}, {ddp_city.strip() or 'City'}, {ddp_country.strip() or 'Country'} (Incoterms 2020)"

# --- 区块 3：Delivery 交期设定 ---
st.markdown("---")
st.subheader("3. 商务条款交期设定 / Commercial Terms Delivery ([x] months)")
delivery_months = st.number_input("交期月数 / Delivery Months ([x])", min_value=1, max_value=60, value=10, step=1, help="用于替换商务条款 Delivery 中的 [x] 变量 / Used to replace [x] in Delivery clause")

# --- 区块 4：W区验证与技术文件手动定价 ---
st.markdown("---")
st.subheader("4. W区验证与技术文件手动定价 / W-Section Manual Pricing (W-2 & W-3)")
col_w2, col_w3 = st.columns(2)
with col_w2:
    w2_price = st.number_input("W-2: FS,HDS,SDS,DS 价格 / Price", min_value=0.0, value=0.0, step=100.0)
with col_w3:
    w3_price = st.number_input("W-3: 第三方仪表校准证书价格 / Third-party Calibration Price", min_value=0.0, value=0.0, step=100.0)

# --- 区块 5：添加特殊文件 (W区自定义附加项) ---
st.markdown("---")
st.subheader("5. 添加特殊文件 / Add Special Documents (W-Section Custom Items - Multi-row)")
default_custom_df = pd.DataFrame([
    {"Model/Description": "", "Technical Configuration": "", "QTY": 1, "Option": "No", "Remark": "", "Unit Price": 0.0}
])
custom_docs_df = st.data_editor(
    default_custom_df, num_rows="dynamic", use_container_width=True,
    key="custom_docs_editor",
    column_config={
        "Option": st.column_config.SelectboxColumn("是否为选配 / Option", options=["No", "Yes"], required=True),
        "Unit Price": st.column_config.NumberColumn("单价 / Unit Price", format="%.2f")
    }
)

# --- 区块 6：V区现场服务设定 ---
st.markdown("---")
st.subheader("6. V区现场服务设定 / V-Section Service Rates & Matrix")
col_rate1, col_rate2 = st.columns(2)
with col_rate1:
    fat_rate = st.number_input("FAT Price Rate (Price / Person / Day)", min_value=0.0, value=500.0, step=50.0)
with col_rate2:
    onsite_rate = st.number_input("Onsite Work Rate (Price / Person / Day)", min_value=0.0, value=600.0, step=50.0)

st.markdown("##### 请填写各项服务的人数与天数 / Please Enter Number of Engineers and Days for Each Service:")
service_items_keys = [
    ("fat", "FAT"),
    ("unpacking", "Unpacking"),
    ("positioning", "Machine Positioning"),
    ("restore", "Restore to Installation"),
    ("commissioning", "Commissioning"),
    ("sat", "SAT"),
    ("iq_oq", "IQ & OQ Execution"),
    ("training", "Training")
]

service_inputs = {}
for key, label in service_items_keys:
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        st.markdown(f"**{label}**")
    with c2:
        eng = st.number_input(f"工程师人数 / Engineers ({label})", min_value=0, value=1, step=1, key=f"eng_{key}")
    with c3:
        days = st.number_input(f"天数 / Days ({label})", min_value=0, value=3, step=1, key=f"days_{key}")
    service_inputs[key] = {"engineer": eng, "days": days}

service_payload = {
    "fat_rate": fat_rate,
    "onsite_rate": onsite_rate,
    "items": service_inputs
}

# --- 区块 7：V区自定义服务项（Packing 之前追加） ---
st.markdown("---")
st.subheader("7. 添加特殊现场服务 / Add Special Onsite Services (Before Packing - Multi-row)")
default_custom_service_df = pd.DataFrame([
    {"Model/Description": "", "Technical Configuration": "", "QTY": 1, "Option": "No", "Remark": "", "Unit Price": 0.0}
])
custom_services_df = st.data_editor(
    default_custom_service_df, num_rows="dynamic", use_container_width=True,
    key="custom_services_editor",
    column_config={
        "Option": st.column_config.SelectboxColumn("是否为选配 / Option", options=["No", "Yes"], required=True),
        "Unit Price": st.column_config.NumberColumn("单价 / Unit Price", format="%.2f")
    }
)

# --- 区块 8：Packing 费用与描述自定义 ---
st.markdown("---")
st.subheader("8. Packing 包装费设定 / Packing Fee Setting")
col_p1, col_p2 = st.columns([2, 1])
with col_p1:
    packing_desc = st.text_input("Packing 描述修改 / Packing Description (Remark Col G)", value="Packed in accordance with international transportation criteria, withstand the conditions of maritime transport.")
with col_p2:
    packing_price = st.number_input("Packing 价格 / Packing Price (Unit Price)", min_value=0.0, value=10000.0, step=500.0)

# --- 区块 9：上传产品部报价单 ---
st.markdown("---")
st.subheader("9. 上传产品部报价单进行转换 / Upload Source Quotation")
uploaded_file = st.file_uploader("支持上传任意产品部 Excel 报价单 / Upload Excel File", type=["xlsx", "xls"])

parsed_machines = None
if uploaded_file is not None:
    try:
        parsed_machines = detect_and_parse_source_excel(uploaded_file)
        st.success(f"✅ 成功解析产品部底价！系统将自动生成 {len(parsed_machines)} 个设备区块。 / Successfully parsed {len(parsed_machines)} equipment block(s).")
    except Exception as e:
        st.error(f"❌ 解析出错 / Parsing error: {e}")

header_data = {
    "customer": customer.strip(), "customer_contacts": customer_contacts.strip(),
    "quotation_no": quotation_no.strip(), "tofflon_contacts": tofflon_contacts.strip(),
    "today_str": today_str, "incoterms_str": final_incoterms_str,
    "currency": currency.strip(), "markup_pct": markup_pct, "exchange_rate": exchange_rate,
    "w2_price": w2_price, "w3_price": w3_price,
    "packing_desc": packing_desc.strip(), "packing_price": packing_price,
    "delivery_months": delivery_months
}

custom_docs_list = []
if custom_docs_df is not None:
    for _, row in custom_docs_df.iterrows():
        if str(row.get("Model/Description", "")).strip():
            custom_docs_list.append(row.to_dict())

custom_services_list = []
if custom_services_df is not None:
    for _, row in custom_services_df.iterrows():
        if str(row.get("Model/Description", "")).strip():
            custom_services_list.append(row.to_dict())

excel_buf = build_quotation_excel(header_data, parsed_machines, custom_docs_list, service_payload, custom_services_list)
st.download_button(
    label="📥 导出全功能标准报价单 Excel / Export Standard Quotation Excel",
    data=excel_buf,
    file_name=f"Quotation_{quotation_no.strip() or 'Test'}_{today_str}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary"
)
