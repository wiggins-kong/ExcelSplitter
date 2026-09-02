#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ExcelSplitter - 按选中列拆分 Excel 工作表为多个文件
（对应 WPS JS 宏 SplitByNonAdjacentSelection 的独立 EXE 实现）

用法:
  GUI 模式 : 直接双击运行（无参数）
  CLI 模式 : ExcelSplitter.exe input.xlsx --cols 1,3 --out 输出目录 --header 1
             （支持 .xlsx / .xlsm / .xls）
"""
import os
import re
import sys
import argparse

import openpyxl

# GUI 依赖（tkinter 在 Windows 系统 Python 中自带；CLI 模式下不需要）
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk, scrolledtext
    _HAS_TK = True
except ImportError:
    _HAS_TK = False

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


# ----------------------------- 界面主题与配色 -----------------------------
COLORS = {
    "bg": "#F4F6F9",
    "surface": "#FFFFFF",
    "primary": "#2F6FED",
    "primary_hover": "#1F5FD6",
    "primary_disabled": "#A9C0F5",
    "text": "#1F2733",
    "text_secondary": "#6B7785",
    "border": "#E2E8F0",
    "success": "#16A34A",
    "danger": "#DC2626",
    "header_bg": "#2F6FED",
}
FONT = ("Microsoft YaHei UI", 10)
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 16, "bold")
FONT_SUB = ("Microsoft YaHei UI", 9)


def _enable_high_dpi():
    """高分屏下避免界面被系统位图拉伸变糊。必须在创建 Tk() 之前调用。"""
    try:
        import ctypes
        # PerMonitorV2（Win10 1703+）：多显示器不同缩放也能清晰
        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
        except Exception:
            pass
        # 回退：按显示器 DPI 感知（Win8.1+）
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass
        # 再回退
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    except Exception:
        pass


def _dpi_scale():
    """返回当前系统 DPI 相对 96 的缩放系数（用于等比放大窗口尺寸）。"""
    try:
        import ctypes
        dpi = ctypes.windll.user32.GetDpiForSystem()
        return max(1.0, dpi / 96.0)
    except Exception:
        return 1.0


def _setup_theme(root):
    if not _HAS_TK:
        raise RuntimeError("tkinter 不可用，无法启动图形界面")
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TFrame", background=COLORS["bg"])
    style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=FONT)
    style.configure("TEntry", font=FONT)
    style.configure("TCombobox", font=FONT)
    style.configure("TSpinbox", font=FONT)
    style.configure("TButton", font=FONT, padding=(10, 6),
                    background=COLORS["surface"], foreground=COLORS["text"],
                    bordercolor=COLORS["border"], relief="solid")
    style.map("TButton",
              background=[("active", "#EEF3FF"), ("disabled", "#F0F0F0")],
              foreground=[("disabled", "#A0A0A0")])
    style.configure("Accent.TButton", font=FONT_BOLD, padding=(12, 8),
                    background=COLORS["primary"], foreground="#FFFFFF", borderwidth=0)
    style.map("Accent.TButton",
              background=[("active", COLORS["primary_hover"]),
                          ("disabled", COLORS["primary_disabled"])])
    style.configure("TProgressbar", thickness=10, borderwidth=0,
                    troughcolor=COLORS["border"], background=COLORS["primary"])
    style.configure("Card.TLabelframe", background=COLORS["surface"],
                    bordercolor=COLORS["border"], relief="solid", borderwidth=1)
    # 滚动条：浅色背景 + 深色滑块，确保可见
    style.configure("Vertical.TScrollbar", background=COLORS["bg"],
                    troughcolor=COLORS["border"], borderwidth=1, relief="flat")
    style.configure("Horizontal.TScrollbar", background=COLORS["bg"],
                    troughcolor=COLORS["border"], borderwidth=1, relief="flat")
    style.configure("Card.TLabelframe.Label", background=COLORS["surface"],
                    foreground=COLORS["text"], font=FONT_BOLD)


def _build_gui():
    if not _HAS_TK:
        print("错误：当前 Python 环境没有 tkinter，无法启动图形界面。")
        print("请使用系统自带的 Python 运行，或直接使用 CLI 模式：")
        print("  python excel_splitter.py 文件.xlsx --cols 1,3 --out 输出目录")
        sys.exit(1)
    _enable_high_dpi()                      # 必须在 Tk() 之前
    root = tk.Tk()
    root.title("Excel 按列拆分工具")
    root.configure(bg=COLORS["bg"])
    _setup_theme(root)
    scale = _dpi_scale()
    root.geometry("%dx%d" % (int(780 * scale), int(780 * scale)))
    root.minsize(int(600 * scale), int(700 * scale))

    path_var = tk.StringVar()
    sheet_var = tk.StringVar()
    header_var = tk.IntVar(value=1)
    out_var = tk.StringVar()
    status_var = tk.StringVar(value="就绪")

    # ---- 顶部标题栏 ----
    header = tk.Frame(root, bg=COLORS["header_bg"], height=int(72 * scale))
    header.pack(fill=tk.X)
    header.pack_propagate(False)
    tk.Label(header, text="Excel 按列拆分工具", bg=COLORS["header_bg"],
             fg="#FFFFFF", font=FONT_TITLE).pack(anchor=tk.W, padx=20, pady=(12, 0))
    tk.Label(header, text="选中列 → 自动按组合拆分 → 每个组合生成一个 Excel 文件",
             bg=COLORS["header_bg"], fg="#DCE7FF", font=FONT_SUB).pack(anchor=tk.W, padx=20, pady=(2, 0))

    # ---- 内容区 ----
    content = ttk.Frame(root, padding=int(16 * scale))
    content.pack(fill=tk.BOTH, expand=True)

    def set_status(msg, kind="info"):
        status_var.set(msg)
        color = {"info": COLORS["text_secondary"], "ok": COLORS["success"],
                 "err": COLORS["danger"]}.get(kind, COLORS["text_secondary"])
        status_label.configure(foreground=color)

    def browse_file():
        p = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")])
        if not p:
            return
        path_var.set(p)
        out_var.set(os.path.join(os.path.dirname(p),
                                os.path.splitext(os.path.basename(p))[0]))
        try:
            sheets = list_sheets(p)
            sheet_combo['values'] = sheets
            sheet_var.set(sheets[0] if sheets else "")
            load_columns()
            set_status("已载入：%s" % os.path.basename(p))
        except Exception as e:
            messagebox.showerror("错误", "无法读取文件:\n" + str(e))

    def load_columns():
        p = path_var.get()
        if not p:
            return
        try:
            hr = header_var.get()
            rows = read_sheet_rows(p, sheet_var.get() or None)
            col_listbox.delete(0, tk.END)
            if not rows or hr < 1 or hr > len(rows):
                return
            header_row = rows[hr - 1]
            for i, name in enumerate(header_row, 1):
                disp = name if (name is not None and str(name).strip() != "") else "(空)"
                col_listbox.insert(tk.END, "第 %d 列: %s" % (i, disp))
        except Exception as e:
            messagebox.showerror("错误", "读取列失败:\n" + str(e))

    def browse_out():
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            out_var.set(d)

    def do_split():
        p = path_var.get()
        if not p:
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return
        sel = list(col_listbox.curselection())
        if not sel:
            messagebox.showwarning("提示",
                "请在「拆分依据列」中选择至少一个列（Ctrl/Shift 可多选）")
            return
        cols = [i + 1 for i in sel]  # listbox 0-based -> 1-based
        out = out_var.get() or os.path.join(
            os.path.dirname(p), os.path.splitext(os.path.basename(p))[0])
        log.delete(1.0, tk.END)
        btn_split.config(state=tk.DISABLED)
        set_status("正在拆分...", "info")
        root.update()
        try:
            def prog(i, t, k):
                pb['value'] = i / t * 100
                set_status("进度 %d/%d  [%s]" % (i, t, k[:40]), "info")
                root.update_idletasks()

            def lg(msg):
                log.insert(tk.END, msg + "\n")
                log.see(tk.END)

            created, errors, out_dir = split_workbook(
                p, sheet=sheet_var.get() or None, col_indices=cols,
                header_row=header_var.get(), out_dir=out,
                progress_cb=prog, log_cb=lg)
            pb['value'] = 100
            if errors:
                set_status("完成：生成 %d 个，失败 %d" % (created, len(errors)), "err")
            else:
                set_status("完成：成功生成 %d 个文件" % created, "ok")
            lg("")
            lg("输出位置: " + out_dir)
            if errors:
                lg("失败 %d 个:" % len(errors))
                for f, e in errors:
                    lg("  %s -> %s" % (f, e))
            messagebox.showinfo("完成",
                "拆分完成！\n生成文件: %d\n失败: %d\n位置: %s"
                % (created, len(errors), out_dir))
        except Exception as e:
            set_status("出错：" + str(e), "err")
            messagebox.showerror("错误", str(e))
        finally:
            btn_split.config(state=tk.NORMAL)

    # ① 选择文件
    f1 = ttk.LabelFrame(content, text="  ① 选择文件  ", padding=int(12 * scale),
                        style="Card.TLabelframe")
    f1.pack(fill=tk.X, pady=(0, int(12 * scale)))
    r1 = ttk.Frame(f1)
    r1.pack(fill=tk.X)
    ttk.Entry(r1, textvariable=path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    ttk.Button(r1, text="浏览...", command=browse_file).pack(side=tk.LEFT)

    # ② 拆分设置
    f2 = ttk.LabelFrame(content, text="  ② 拆分设置  ", padding=int(12 * scale),
                        style="Card.TLabelframe")
    f2.pack(fill=tk.X, pady=(0, int(12 * scale)))
    opt = ttk.Frame(f2)
    opt.pack(fill=tk.X, pady=(0, 8))
    ttk.Label(opt, text="工作表").pack(side=tk.LEFT)
    sheet_combo = ttk.Combobox(opt, textvariable=sheet_var, state="readonly",
                               width=20, font=FONT)
    sheet_combo.pack(side=tk.LEFT, padx=(6, 16))
    sheet_combo.bind("<<ComboboxSelected>>", lambda e: load_columns())
    ttk.Label(opt, text="表头行").pack(side=tk.LEFT)
    ttk.Spinbox(opt, from_=1, to=50, textvariable=header_var, width=6,
                font=FONT, command=load_columns).pack(side=tk.LEFT, padx=(6, 0))
    ttk.Label(f2, text="拆分依据列（Ctrl / Shift 可多选）").pack(anchor=tk.W, pady=(0, 4))
    col_listbox = tk.Listbox(f2, selectmode=tk.EXTENDED, height=8, font=FONT,
                             bg=COLORS["surface"], fg=COLORS["text"],
                             selectbackground="#D6E4FF", selectforeground=COLORS["text"],
                             selectborderwidth=0,
                             relief="solid", borderwidth=1, highlightthickness=0,
                             exportselection=False)
    col_listbox.pack(fill=tk.X)

    # ③ 输出位置
    f3 = ttk.LabelFrame(content, text="  ③ 输出位置  ", padding=int(12 * scale),
                        style="Card.TLabelframe")
    f3.pack(fill=tk.X, pady=(0, int(12 * scale)))
    r3 = ttk.Frame(f3)
    r3.pack(fill=tk.X)
    ttk.Entry(r3, textvariable=out_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
    ttk.Button(r3, text="浏览...", command=browse_out).pack(side=tk.LEFT)

    # 主操作
    btn_split = ttk.Button(content, text="开始拆分", command=do_split, style="Accent.TButton")
    btn_split.pack(fill=tk.X, pady=(0, int(12 * scale)), ipady=int(6 * scale))

    # 状态 + 进度
    status_label = ttk.Label(content, textvariable=status_var,
                             foreground=COLORS["text_secondary"])
    status_label.pack(anchor=tk.W, pady=(0, 4))
    pb = ttk.Progressbar(content, orient=tk.HORIZONTAL, mode="determinate", maximum=100)
    pb.pack(fill=tk.X, pady=(0, int(12 * scale)))

    # 运行日志（带滚动条，保证最小高度 120px）
    f4 = ttk.LabelFrame(content, text="  运行日志  ", padding=int(8 * scale),
                        style="Card.TLabelframe")
    f4.pack(fill=tk.BOTH, expand=True, pady=(0, int(4 * scale)))
    f4.pack_propagate(False)
    log = scrolledtext.ScrolledText(f4, font=FONT, bg=COLORS["surface"], fg=COLORS["text"],
                                    wrap=tk.WORD,
                                    relief="solid", borderwidth=1, highlightthickness=0,
                                    height=int(6 * scale))
    log.pack(fill=tk.BOTH, expand=True)

    root.mainloop()


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
            _build_gui()
        except Exception as e:
            print("无法启动图形界面: " + str(e))
            sys.exit(1)


if __name__ == "__main__":
    main()
