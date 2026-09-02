#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""ExcelSplitter GUI —— sv_ttk 新界面（v1.1.0 重构版）

入口：被 excel_splitter.py（无参数）导入调用；也可单独运行：
  python excel_splitter_gui.py

依赖：tkinter（系统 Python 自带）、sv_ttk、tkinterdnd2（拖拽，可选降级）
注意：Tk 9.0 + sv_ttk 会在 update_idletasks 后用主题默认值重置 tk 部件的
      -background，因此本界面所有自定义颜色一律走 ttk.Style，不用 tk 部件显式设色。
"""
import ctypes
import json
import os
import queue
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 高 DPI：必须在 Tk() 创建之前声明进程感知，否则高分屏下界面被系统位图拉伸变糊
try:
    ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)   # PerMonitorV2
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

import sv_ttk

try:  # 拖拽支持（环境缺失时降级为普通窗口，仅提示不可用）
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:
    _HAS_DND = False

import excel_splitter as core

# ----------------------------- 设计令牌 -----------------------------
ACCENT = "#2563eb"
TEXT = "#1f2937"
SUB = "#6b7280"
OK = "#16a34a"
DANGER = "#dc2626"
WHITE = "#ffffff"          # 卡片底色（sv_ttk Labelframe 近白，重置无感）
BG90 = "#fafafa"           # sv_ttk light 窗口底色

FONT = ("Microsoft YaHei UI", 10)
FONT_SM = ("Microsoft YaHei UI", 9)
FONT_H1 = ("Microsoft YaHei UI", 16, "bold")
FONT_H2 = ("Microsoft YaHei UI", 11, "bold")
FONT_BADGE = ("Microsoft YaHei UI", 10, "bold")
FONT_LOGF = ("Consolas", 9)


def _settings_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ExcelSplitter", "settings.json")


def load_settings():
    """读取设置（只记输出目录等轻量字段），文件不存在/损坏返回 {}。"""
    try:
        with open(_settings_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(out_dir):
    try:
        p = _settings_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump({"out_dir": out_dir}, f, ensure_ascii=False)
    except Exception:
        pass


# ----------------------------- 工具构件 -----------------------------
def _rounded_badge(parent, num):
    """accent 色圆角数字徽章（Canvas 背景取近白，被主题重置也无感）。"""
    c = tk.Canvas(parent, width=30, height=30, bg=WHITE, highlightthickness=0)
    c.pack(side=tk.LEFT, padx=(0, 10))
    r = 8
    c.create_polygon(
        2 + r, 2, 28 - r, 2, 28, 2, 28, 2 + r, 28, 28 - r, 28, 28,
        28 - r, 28, 2 + r, 28, 2, 28, 2, 28 - r, 2, 2 + r, 2, 2,
        smooth=True, fill=ACCENT, outline="")
    c.create_text(15, 15, text=num, fill="white", font=FONT_BADGE)
    return c


class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.path_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.header_var = tk.IntVar(value=1)
        self.out_var = tk.StringVar()
        self.status_var = tk.StringVar(value="就绪 · 选择文件开始")
        self.rows_cache = []           # 当前文件数据行缓存（含表头），避免重复读盘
        self.hint_var = tk.StringVar(value="选择拆分列后显示预计生成数")
        self._polling = False

        s = load_settings()
        if s.get("out_dir"):
            self.out_var.set(s["out_dir"])

        self._setup_styles()
        self._build_header()
        self._build_cards()
        self._build_op_area()
        self._build_log()

        if _HAS_DND:
            root.drop_target_register(DND_FILES)
            root.dnd_bind("<<Drop>>", self._on_drop)
        else:
            self.log("提示: 拖拽组件不可用（tkinterdnd2 缺失），请用「浏览…」选择文件")

        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- 样式（全部走 ttk.Style，绕开 Tk 9.0 背景重置坑） ----------
    def _setup_styles(self):
        style = ttk.Style()
        style.configure("Header.TFrame", background=ACCENT)
        style.configure("HeaderTitle.TLabel", background=ACCENT, foreground="#ffffff", font=FONT_H1)
        style.configure("HeaderSub.TLabel", background=ACCENT, foreground="#dbe6ff", font=FONT_SM)
        style.configure("CardTitle.TLabel", background=WHITE, foreground=TEXT, font=FONT_H2)
        style.configure("CardHint.TLabel", background=WHITE, foreground=SUB, font=FONT_SM)
        style.configure("StatusOK.TLabel", background=BG90, foreground=OK, font=FONT_SM)
        style.configure("StatusErr.TLabel", background=BG90, foreground=DANGER, font=FONT_SM)
        style.configure("StatusInfo.TLabel", background=BG90, foreground=SUB, font=FONT_SM)
        style.configure("Treeview", rowheight=26, font=FONT)
        style.configure("Treeview.Heading", font=FONT)

    # ---------- 布局 ----------
    def _build_header(self):
        header = ttk.Frame(self.root, style="Header.TFrame")
        header.pack(fill=tk.X)
        ttk.Label(header, text="Excel 按列拆分工具",
                  style="HeaderTitle.TLabel").pack(anchor=tk.W, padx=24, pady=(14, 0))
        ttk.Label(header, text="选择拆分列 → 自动按组合拆分 → 每个组合生成一个独立 Excel 文件",
                  style="HeaderSub.TLabel").pack(anchor=tk.W, padx=24, pady=(2, 12))

    def _card(self, num, title, hint="", parent=None):
        f = ttk.LabelFrame(parent or self.content, padding=(16, 12, 16, 14))
        f.pack(fill=tk.X, pady=(0, 12))
        head = ttk.Frame(f)
        head.pack(fill=tk.X, pady=(0, 10))
        _rounded_badge(head, num)
        ttk.Label(head, text=title, style="CardTitle.TLabel").pack(side=tk.LEFT, anchor=tk.S)
        if hint:
            ttk.Label(head, text=hint, style="CardHint.TLabel").pack(side=tk.RIGHT, anchor=tk.S)
        return f

    def _build_cards(self):
        self.content = ttk.Frame(self.root, padding=(18, 12))
        self.content.pack(fill=tk.BOTH, expand=True)
        # 底部通栏(状态+进度)
        self.bottom = ttk.Frame(self.content)
        self.bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        # 主区两栏
        main = ttk.Frame(self.content)
        main.pack(fill=tk.BOTH, expand=True)
        self.content.columnconfigure(0, weight=1)
        main.columnconfigure(0, weight=3, uniform="col")
        main.columnconfigure(1, weight=2, uniform="col")
        main.rowconfigure(0, weight=1)
        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        self._log_parent = right
        left.rowconfigure(1, weight=1)
        right.rowconfigure(1, weight=1)

        # ── 左列:① 选择文件 ──
        c1 = self._card("1", "选择文件", "支持 .xlsx / .xlsm / .xls", parent=left)
        r1 = ttk.Frame(c1)
        r1.pack(fill=tk.X)
        ttk.Entry(r1, textvariable=self.path_var, font=FONT).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(r1, text="浏览…", command=self.browse_file).pack(side=tk.LEFT)
        ttk.Label(c1, text="也可以把 Excel 文件直接拖到窗口上",
                  style="CardHint.TLabel").pack(anchor=tk.W, pady=(8, 0))

        # ── 左列:② 拆分设置(弹性) ──
        c2 = self._card("2", "拆分设置", parent=left)
        c2.pack(fill=tk.BOTH, expand=True)
        opt = ttk.Frame(c2)
        opt.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(opt, text="工作表").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(opt, textvariable=self.sheet_var, state="readonly",
                                        width=14, font=FONT)
        self.sheet_combo.pack(side=tk.LEFT, padx=(6, 18))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.load_columns())
        ttk.Label(opt, text="表头行").pack(side=tk.LEFT)
        self.header_spin = ttk.Spinbox(opt, from_=1, to=50, textvariable=self.header_var,
                                       width=5, font=FONT, command=self.load_columns)
        self.header_spin.pack(side=tk.LEFT, padx=(6, 0))
        self.header_spin.bind("<KeyRelease>", lambda e: self.load_columns())

        ttk.Label(c2, text="拆分依据列（按住 Ctrl / Shift 可多选）").pack(anchor=tk.W, pady=(0, 6))
        tw = ttk.Frame(c2)
        tw.pack(fill=tk.BOTH, expand=True)
        self.col_tree = ttk.Treeview(tw, columns=("no", "name", "preview"), show="headings",
                                     height=6, selectmode="extended")
        self.col_tree.heading("no", text="#")
        self.col_tree.heading("name", text="列名")
        self.col_tree.heading("preview", text="示例值")
        self.col_tree.column("no", width=42, anchor=tk.CENTER, stretch=False)
        self.col_tree.column("name", width=180)
        self.col_tree.column("preview", width=210)
        vsb = ttk.Scrollbar(tw, orient=tk.VERTICAL, command=self.col_tree.yview)
        self.col_tree.configure(yscrollcommand=vsb.set)
        self.col_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.col_tree.bind("<<TreeviewSelect>>", lambda e: self.update_estimate())

        # ── 左列底部:开始按钮 ──
        self.btn_split = ttk.Button(left, text="开始拆分", style="Accent.TButton",
                                    command=self.do_split)
        self.btn_split.pack(fill=tk.X, ipady=5)

        # ── 右列:③ 输出位置 ──
        c3 = self._card("3", "输出位置", parent=right)
        self.hint_label = ttk.Label(c3, textvariable=self.hint_var, style="CardHint.TLabel")
        self.hint_label.pack(side=tk.TOP, anchor=tk.E, pady=(0, 4))
        r3 = ttk.Frame(c3)
        r3.pack(fill=tk.X)
        ttk.Entry(r3, textvariable=self.out_var, font=FONT).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(r3, text="浏览…", command=self.browse_out).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(r3, text="打开输出目录", command=self.open_out_dir).pack(side=tk.LEFT)

    def _build_op_area(self):
        """底部通栏:状态指示 + 进度条。"""
        self.status_dot = ttk.Label(self.bottom, text="●", style="StatusInfo.TLabel")
        self.status_dot.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(self.bottom, textvariable=self.status_var, style="StatusInfo.TLabel").pack(side=tk.LEFT)
        self.pb = ttk.Progressbar(self.bottom, mode="determinate", maximum=100)
        self.pb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(14, 0))

    def _build_log(self):
        c4 = self._card("4", "运行日志", parent=self._log_parent)
        c4.pack(fill=tk.BOTH, expand=True)
        lw = ttk.Frame(c4)
        lw.pack(fill=tk.BOTH, expand=True)
        self.log_tree = ttk.Treeview(lw, columns=("time", "msg"), show="headings", height=6)
        self.log_tree.heading("time", text="时间")
        self.log_tree.heading("msg", text="消息")
        self.log_tree.column("time", width=80, anchor=tk.W, stretch=False)
        self.log_tree.column("msg", width=280)
        vsb = ttk.Scrollbar(lw, orient=tk.VERTICAL, command=self.log_tree.yview)
        self.log_tree.configure(yscrollcommand=vsb.set)
        self.log_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_tree.tag_configure("ok", foreground=OK)
        self.log_tree.tag_configure("err", foreground=DANGER)

    # ---------- 行为 ----------
    def log(self, msg, tag=None):
        ts = time_str()
        iid = self.log_tree.insert("", tk.END, values=(ts, msg), tags=(tag,) if tag else ())
        self.log_tree.see(iid)
        # 日志过长时裁剪，避免行数无限增长
        kids = self.log_tree.get_children()
        if len(kids) > 500:
            self.log_tree.delete(kids[0])

    def set_status(self, msg, kind="info"):
        self.status_var.set(msg)
        self.status_dot.configure(style={
            "info": "StatusInfo.TLabel", "ok": "StatusOK.TLabel",
            "err": "StatusErr.TLabel"}.get(kind, "StatusInfo.TLabel"))

    def browse_file(self):
        p = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")])
        if p:
            self.load_file(p)

    def verify_path(self, p):
        if not p:
            return False
        if not os.path.isfile(p):
            messagebox.showerror("错误", "文件不存在:\n" + p)
            return False
        if os.path.splitext(p)[1].lower() not in (".xlsx", ".xlsm", ".xls"):
            messagebox.showerror("错误", "不支持的格式，仅支持 .xlsx / .xlsm / .xls")
            return False
        return True

    def load_file(self, p):
        """载入文件：缓存数据行、刷工作表下拉与列列表。"""
        if not self.verify_path(p):
            return
        try:
            sheets = core.list_sheets(p)
            sheet = self.sheet_var.get()
            if sheet not in sheets:
                sheet = sheets[0] if sheets else ""
                self.sheet_var.set(sheet)
            self.sheet_combo["values"] = sheets
            self.rows_cache = core.read_sheet_rows(p, sheet or None)
            self.path_var.set(p)
            if not self.out_var.get():  # 无记忆时默认输出到源文件同名目录
                self.out_var.set(os.path.join(os.path.dirname(p),
                                              os.path.splitext(os.path.basename(p))[0]))
            self.load_columns()
            self.log("已载入: " + os.path.basename(p))
            self.set_status("已载入 %s · %d 行数据" % (
                os.path.basename(p), len(self.rows_cache)), "ok")
        except Exception as e:
            messagebox.showerror("错误", "无法读取文件:\n" + str(e))

    def load_columns(self):
        tree = self.col_tree
        tree.delete(*tree.get_children())
        rows = self.rows_cache
        if not rows:
            return
        hr = self.header_var.get()
        if hr < 1 or hr > len(rows):
            return
        header = rows[hr - 1]
        data = rows[hr:]
        ncols = max(len(header), 1)
        for i in range(1, ncols + 1):
            name = header[i - 1] if i - 1 < len(header) else None
            disp = name if (name is not None and str(name).strip() != "") else "(空)"
            tree.insert("", tk.END, iid=str(i), values=(i, disp, self._sample(data, i - 1)))
        self.update_estimate()

    def _sample(self, data, col, maxn=2):
        """取该列前几个非空值拼接示例。"""
        vals = []
        for row in data:
            if col < len(row) and row[col] is not None and str(row[col]).strip() != "":
                s = str(row[col]).strip()
                if len(s) > 18:
                    s = s[:18] + "…"
                vals.append(s)
                if len(vals) >= maxn:
                    break
        return " | ".join(vals) + (" …" if vals else "")

    def selected_cols(self):
        return [int(i) for i in self.col_tree.selection()]

    def update_estimate(self):
        """实时估算：所选列组合将生成多少个文件。"""
        cols = self.selected_cols()
        rows = self.rows_cache
        hr = self.header_var.get()
        if not cols or not rows or hr < 1 or hr > len(rows):
            self.hint_var.set("选择拆分列后显示预计生成数")
            return
        try:
            _, _, order = core.group_keys(rows, cols, hr)
            n = len(order)
            self.hint_var.set("预计生成 %d 个文件" % n)
        except Exception:
            self.hint_var.set("")

    def browse_out(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_var.set(d)

    def open_out_dir(self):
        d = self.out_var.get()
        if not d or not os.path.isdir(d):
            messagebox.showinfo("提示", "输出目录还不存在，拆分完成后会自动创建。")
            return
        try:
            os.startfile(d)
        except Exception as e:
            messagebox.showerror("错误", "无法打开目录:\n" + str(e))

    def _on_drop(self, event):
        raw = event.data.strip()
        files = re.findall(r"\{([^{}]+)\}", raw) or raw.split()
        for f in files:
            if os.path.isfile(f) and os.path.splitext(f)[1].lower() in (".xlsx", ".xlsm", ".xls"):
                self.load_file(f)
                return
        messagebox.showerror("错误", "拖入的不是有效的 Excel 文件")

    def do_split(self):
        if self._polling:
            return
        p = self.path_var.get()
        if not p:
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return
        cols = self.selected_cols()
        if not cols:
            messagebox.showwarning("提示", "请在「拆分依据列」中选择至少一个列（Ctrl/Shift 可多选）")
            return
        out = self.out_var.get()
        hr = self.header_var.get()
        save_settings(out)  # 顺手记忆输出目录

        self.btn_split.configure(state=tk.DISABLED)
        self.pb["value"] = 0
        self.log_tree.delete(*self.log_tree.get_children())
        self.set_status("正在拆分…", "info")
        q = queue.Queue()
        self._polling = True

        def worker():
            try:
                created, errors, out_dir = core.split_workbook(
                    p, sheet=self.sheet_var.get() or None, col_indices=cols,
                    header_row=hr, out_dir=out,
                    progress_cb=lambda i, t, k: q.put(("prog", i, t, k)),
                    log_cb=lambda m: q.put(("log", m)))
                q.put(("done", created, errors, out_dir))
            except Exception as e:
                q.put(("error", str(e)))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, lambda: self._poll(q))

    def _poll(self, q):
        try:
            while True:
                item = q.get_nowait()
                kind = item[0]
                if kind == "prog":
                    _, i, t, k = item
                    self.pb["value"] = i / t * 100
                    self.set_status("进度 %d/%d · %s" % (i, t, k[:32]), "info")
                elif kind == "log":
                    self.log(item[1])
                elif kind == "done":
                    _, created, errors, out_dir = item
                    self._finish(created, errors, out_dir)
                    return
                elif kind == "error":
                    self._polling = False
                    self.btn_split.configure(state=tk.NORMAL)
                    self.set_status("出错：" + item[1], "err")
                    messagebox.showerror("错误", item[1])
                    return
        except queue.Empty:
            pass
        if self._polling:
            self.root.after(100, lambda: self._poll(q))

    def _finish(self, created, errors, out_dir):
        self._polling = False
        self.pb["value"] = 100
        self.btn_split.configure(state=tk.NORMAL)
        if errors:
            self.set_status("完成：生成 %d 个，失败 %d" % (created, len(errors)), "err")
        else:
            self.set_status("完成：成功生成 %d 个文件" % created, "ok")
        self.log("输出位置: " + out_dir, "ok")
        if errors:
            self.log("失败 %d 个：" % len(errors), "err")
            for f, e in errors:
                self.log("  %s -> %s" % (f, e), "err")
        messagebox.showinfo("完成", "拆分完成！\n生成文件: %d\n失败: %d\n位置: %s"
                            % (created, len(errors), out_dir))

    def _on_close(self):
        save_settings(self.out_var.get())
        self.root.destroy()


def time_str():
    import time
    return time.strftime("%H:%M:%S")


def run_gui():
    """GUI 入口（由 excel_splitter.py 无参数调用，或本文件直接运行）。"""
    if _HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    root.title("Excel 按列拆分工具")
    root.geometry("980x680+120+60")
    root.minsize(860, 600)
    sv_ttk.set_theme("light")
    ExcelSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()