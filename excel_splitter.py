#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ExcelSplitter - 按选中列拆分 Excel 工作表为多个文件
（对应 WPS JS 宏 SplitByNonAdjacentSelection 的独立 EXE 实现）

用法:
  GUI 模式 : 直接双击运行（无参数，需 tkinter + sv_ttk + tkinterdnd2）
  CLI 模式 : ExcelSplitter.exe input.xlsx --cols 1,3 --out 输出目录 --header 1
             （支持 .xlsx / .xlsm / .xls）

文件分工:
  excel_splitter.py        核心逻辑 + CLI 入口（本文件，无 GUI 依赖）
  excel_splitter_gui.py    sv_ttk 图形界面（GUI 模式入口，import 本文件）
"""
import os
import re
import sys
import argparse

import openpyxl

# .xls 旧格式读取依赖（openpyxl 只支持 .xlsx / .xlsm）
try:
    import xlrd
    _HAS_XLRD = True
except ImportError:
    _HAS_XLRD = False

# Windows 文件名非法字符
_ILLEGAL = re.compile(r'[\\/:*?"<>|]')


def _is_xls(path):
    """按扩展名判断是否为 .xls 旧格式。"""
    return os.path.splitext(path)[1].lower() == ".xls"


def _require_xlrd():
    if not _HAS_XLRD:
        raise ValueError("读取 .xls 文件需要 xlrd 库，请先安装：pip install xlrd")


def _xls_value(sh, r, c):
    """xlrd 单元格取值，日期类型转成 datetime（与 openpyxl 行为对齐）。"""
    if sh.cell_type(r, c) == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate_as_datetime(sh.cell_value(r, c), sh.book.datemode)
        except Exception:
            pass
    return sh.cell_value(r, c)


def list_sheets(path):
    """返回工作簿内所有工作表名（按扩展名自动分派 .xls / .xlsx）。"""
    if _is_xls(path):
        _require_xlrd()
        wb = xlrd.open_workbook(path)
        try:
            return wb.sheet_names()
        finally:
            wb.release_resources()
    wb = openpyxl.load_workbook(path)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def read_sheet_rows(path, sheet=None):
    """读取指定工作表为 rows（list[list]，含表头行）。按扩展名自动分派。

    注意：xlsx 不能用 read_only=True！部分工具导出的 xlsx 内部 <dimension>
    声明错误（如只写 A1，实际数据到 AK2629），read_only 模式信任该声明会把
    工作表截断成 1x1，导致只读到第一个表头单元格。常规模式不受影响。
    """
    if _is_xls(path):
        _require_xlrd()
        wb = xlrd.open_workbook(path)
        try:
            sh = wb.sheet_by_name(sheet) if sheet else wb.sheet_by_index(0)
            return [[_xls_value(sh, r, c) for c in range(sh.ncols)]
                    for r in range(sh.nrows)]
        finally:
            wb.release_resources()
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb[sheet] if sheet else wb.active
        return [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def clean_cell(value):
    """清洗单元格值，模仿原宏的强力清洗逻辑：
    空值/None -> '空白'；去首尾空格；去单元格内换行/制表符。"""
    if value is None or value == "":
        return "空白"
    s = str(value).strip()
    s = re.sub(r'[\r\n\t]+', ' ', s)
    if s == "":
        return "空白"
    return s


def safe_name(key):
    """根据分组键生成安全的文件名（处理非法字符与长度）。"""
    name = _ILLEGAL.sub('-', key)
    name = name.replace(':', '：')
    if len(name) > 150:
        name = name[:150]
    return name


def autofit(ws, max_width=60):
    """粗略自适应列宽（CJK 字符按 2 计宽）。"""
    for col in ws.columns:
        letter = None
        length = 0
        for cell in col:
            if letter is None:
                letter = cell.column_letter
            v = cell.value
            if v is not None:
                s = str(v)
                w = sum(2 if ord(ch) > 0x2E80 else 1 for ch in s)
                if w > length:
                    length = w
        if letter:
            ws.column_dimensions[letter].width = min(max(length + 2, 8), max_width)


def group_keys(rows, col_indices, header_row=1):
    """只统计分组结果（不写文件），供 GUI 实时显示"预计生成 N 个文件"。
    返回 (data_rows, groups_keys:list)。与 split_workbook 的分组口径完全一致。
    """
    col_indices = sorted(set(int(i) for i in col_indices))
    sel = [i - 1 for i in col_indices]
    header = list(rows[header_row - 1])
    data_rows = rows[header_row:]
    groups = {}
    order = []
    max_idx = max(sel)
    for row in data_rows:
        row = list(row)
        if max_idx >= len(row):
            row = row + [None] * (max_idx - len(row) + 1)
        parts = [clean_cell(row[c]) for c in sel]
        key = tuple(parts)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)
    return header, data_rows, order


def split_workbook(path, sheet=None, col_indices=None, header_row=1,
                   out_dir=None, progress_cb=None, log_cb=None):
    """
    按 col_indices(1-based 列表) 组合拆分工作表。
    每个唯一组合生成一个 .xlsx（含完整表头与所有列）。
    返回 (created_count, errors_list, out_dir)。
    """
    if not col_indices:
        raise ValueError("未指定拆分列")
    col_indices = sorted(set(int(i) for i in col_indices))
    sel = [i - 1 for i in col_indices]  # 转 0-based

    # 统一读取：.xls 走 xlrd、.xlsx/.xlsm 走 openpyxl 常规模式
    # （不能用 read_only=True，原因见 read_sheet_rows 注释）
    rows = read_sheet_rows(path, sheet)

    if not rows:
        raise ValueError("工作表为空")
    if header_row < 1 or header_row > len(rows):
        raise ValueError("表头行超出范围")

    header = list(rows[header_row - 1])
    data_rows = rows[header_row:]
    if not data_rows:
        raise ValueError("没有数据行（表头行之后为空）")

    # 分组
    groups = {}
    order = []
    max_idx = max(sel)
    for row in data_rows:
        row = list(row)
        if max_idx >= len(row):
            row = row + [None] * (max_idx - len(row) + 1)
        parts = [clean_cell(row[c]) for c in sel]
        key = tuple(parts)  # 用元组作分组键，避免分隔符与数据冲突
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    if out_dir is None:
        base = os.path.splitext(os.path.basename(path))[0]
        out_dir = os.path.join(os.path.dirname(os.path.abspath(path)), base)
    os.makedirs(out_dir, exist_ok=True)

    created = 0
    errors = []
    total = len(order)
    for idx, key in enumerate(order, 1):
        disp = " - ".join(key)
        if progress_cb:
            progress_cb(idx, total, disp)
        fname = safe_name(disp) + ".xlsx"
        full = os.path.join(out_dir, fname)
        try:
            nwb = openpyxl.Workbook()
            nws = nwb.active
            nws.append(list(header))
            for r in groups[key]:
                nws.append(r)
            autofit(nws)
            nwb.save(full)
            nwb.close()
            created += 1
            if log_cb:
                log_cb("已生成: " + fname)
        except Exception as e:
            errors.append((fname, str(e)))
            if log_cb:
                log_cb("失败: " + fname + " -> " + str(e))
    return created, errors, out_dir


def main():
    parser = argparse.ArgumentParser(description="按选中列拆分 Excel")
    parser.add_argument("input", nargs="?", help="输入 xlsx/xls 路径")
    parser.add_argument("--cols", help="拆分列(1-based,逗号分隔), 如 1,3")
    parser.add_argument("--sheet", help="工作表名")
    parser.add_argument("--header", type=int, default=1, help="表头行(默认1)")
    parser.add_argument("--out", help="输出文件夹")
    args = parser.parse_args()

    if args.input and args.cols:
        cols = [int(x.strip()) for x in args.cols.split(",") if x.strip()]
        created, errors, out_dir = split_workbook(
            args.input, sheet=args.sheet, col_indices=cols,
            header_row=args.header, out_dir=args.out,
            progress_cb=lambda i, t, k: print("进度 %d/%d %s" % (i, t, k[:30])),
            log_cb=lambda m: print(m))
        print("\n完成：生成 %d 个文件，失败 %d" % (created, len(errors)))
        print("位置: " + out_dir)
        for f, e in errors:
            print("  失败 %s -> %s" % (f, e))
    else:
        try:
            from excel_splitter_gui import run_gui
            run_gui()
        except Exception as e:
            print("无法启动图形界面: " + str(e))
            print("提示：GUI 模式需要 tkinter / sv_ttk / tkinterdnd2。")
            print("可改用 CLI 模式：python excel_splitter.py 文件.xlsx --cols 1,3 --out 输出目录")
            sys.exit(1)


if __name__ == "__main__":
    main()