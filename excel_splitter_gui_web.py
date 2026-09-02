#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""ExcelSplitter GUI (v2.0.0) —— pywebview + WebView2 + 真 Mica 实现

入口：被 excel_splitter.py（无参数）导入调用；也可单独运行：
  python excel_splitter_gui_web.py

依赖：pywebview（Windows 上使用系统自带的 WebView2 运行时）
Mica：Win11 22H2+ 通过 DwmSetWindowAttribute 开启系统云母背景；
      旧系统自动回退为 CSS 模拟（前端 body.no-mica）。
"""
import ctypes
import json
import os
import queue
import sys
import threading

import webview

import excel_splitter as core

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GUI_DIR = os.path.join(BASE_DIR, "webgui")
WINDOW_TITLE = "ExcelSplitter"

# Win11 Mica 常量
_DWMWA_USE_IMMERSIVE_DARK_MODE = 20
_DWMWA_SYSTEMBACKDROP_TYPE = 38          # 22H2+ (build 22621)
_DWMWA_MICA_EFFECT = 1029                # 21H2 insider / 22000
_DWMSBT_MAINWINDOW = 2                   # Mica


def _settings_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ExcelSplitter", "settings.json")


def load_settings():
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


def _clean_preview(v, maxlen=18):
    s = str(v).strip() if v is not None and str(v).strip() != "" else ""
    if len(s) > maxlen:
        s = s[:maxlen] + "…"
    return s


class Api:
    """暴露给前端 JS 的桥（pywebview js_api）。方法名对应 pywebview.api.xxx"""

    def __init__(self):
        self._win = None
        self._splitting = False
        self._dom_ok = False
        self._mica_ok = False

    # ---------- 内部 ----------
    def _emit(self, js):
        """向后端窗口推送 JS 事件（线程安全）。"""
        if self._win:
            try:
                self._win.evaluate_js(js)
            except Exception:
                pass

    def _load_data(self, path, sheet=None, header_row=1):
        sheets = core.list_sheets(path)
        if sheet not in sheets:
            sheet = sheets[0] if sheets else None
        rows = core.read_sheet_rows(path, sheet)
        header_row = max(1, min(header_row, len(rows)))
        header = rows[header_row - 1]
        data = rows[header_row:]
        ncols = max(len(header), 1)
        columns = []
        for i in range(1, ncols + 1):
            name = header[i - 1] if i - 1 < len(header) else None
            disp = str(name) if (name is not None and str(name).strip() != "") else "(空)"
            previews = []
            for row in data:
                if i - 1 < len(row) and row[i - 1] is not None and str(row[i - 1]).strip() != "":
                    previews.append(_clean_preview(row[i - 1]))
                    if len(previews) >= 2:
                        break
            columns.append({"no": i, "name": disp, "preview": " | ".join(previews) + (" …" if previews else "")})
        return {"sheets": sheets, "sheet": sheet, "columns": columns, "rows": len(data)}

    # ---------- JS API ----------
    def get_settings(self):
        return load_settings()

    def mica_active(self):
        """当前窗口是否成功开启真 Mica（前端据此决定是否启用 CSS 回退）。"""
        return getattr(self, "_mica_ok", False)

    def pick_file(self):
        result = self._win.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel 文件 (*.xlsx;*.xlsm;*.xls)", "所有文件 (*.*)"),
        )
        if not result:
            return None
        path = result[0] if isinstance(result, (list, tuple)) else result
        return self.load_path(path)

    def load_path(self, path):
        try:
            if not path:
                return {"error": "未获取到文件路径"}
            ext = os.path.splitext(path)[1].lower()
            if ext not in (".xlsx", ".xlsm", ".xls"):
                return {"error": "不支持的格式，仅支持 .xlsx / .xlsm / .xls"}
            if not os.path.isfile(path):
                return {"error": "文件不存在：" + path}
            info = self._load_data(path)
            out_dir = load_settings().get("out_dir", "")
            if not out_dir:
                base = os.path.splitext(os.path.basename(path))[0]
                out_dir = os.path.join(os.path.dirname(path), base)
            return dict(path=path, out_dir=out_dir, **info)
        except Exception as e:
            return {"error": "无法读取文件：" + str(e)}

    def load_blob(self, filename, data_b64):
        """WebView2 拖放不暴露真实路径：前端读文件内容（base64）传到这里，
        落到临时目录后按普通文件载入。"""
        import base64
        import tempfile
        try:
            safe_name = os.path.basename(filename or "dropped.xlsx")
            ext = os.path.splitext(safe_name)[1].lower()
            if ext not in (".xlsx", ".xlsm", ".xls"):
                return {"error": "不支持的格式，仅支持 .xlsx / .xlsm / .xls"}
            tmp_dir = os.path.join(tempfile.gettempdir(), "ExcelSplitter")
            os.makedirs(tmp_dir, exist_ok=True)
            path = os.path.join(tmp_dir, safe_name)
            with open(path, "wb") as f:
                f.write(base64.b64decode(data_b64))
            return self.load_path(path)
        except Exception as e:
            return {"error": "拖入文件读取失败：" + str(e)}

    def get_columns(self, path, sheet, header_row):
        try:
            info = self._load_data(path, sheet or None, int(header_row))
            return {"columns": info["columns"], "rows": info["rows"]}
        except Exception as e:
            return {"error": str(e)}

    def estimate(self, path, sheet, header_row, cols):
        try:
            rows = core.read_sheet_rows(path, sheet or None)
            _, _, order = core.group_keys(rows, [int(c) for c in cols], int(header_row))
            return {"count": len(order)}
        except Exception:
            return {"count": None}

    def pick_out_dir(self, current):
        result = self._win.create_file_dialog(webview.FOLDER_DIALOG)
        if not result:
            return None
        return result[0] if isinstance(result, (list, tuple)) else result

    def set_out_dir(self, path):
        """兜底通道：校验路径是目录后设为输出目录并记忆。"""
        try:
            if path and os.path.isdir(path):
                save_settings(path)
                return {"ok": True, "path": path}
            return {"ok": False}
        except Exception:
            return {"ok": False}

    # ---------- 原生 DOM 拖放（可拿真实路径，支持文件夹） ----------
    def _drop_log(self, msg):
        """诊断日志：写入 %TEMP%\\ExcelSplitter\\drop.log（排查拖放问题用）。"""
        try:
            import tempfile
            d = os.path.join(tempfile.gettempdir(), "ExcelSplitter")
            os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, "drop.log"), "a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            pass

    def _on_dom_dragover(self, e):
        """放行拖入（prevent 默认行为，否则 Explorer 不显示可放置）。"""

    def _on_dom_drop(self, e):
        """window.dom 原生 drop 事件：按当前 hover 区域路由到文件区或文件夹区。

        前端拖放区 handler 不再 stopPropagation / 提前清空 __hoverZone，
        drop 事件必须冒泡到 document 才能触发这里的原生监听（从而拿到
        pywebviewFullPath 真实路径）。zone 在本方法内「消费」式读取。"""
        try:
            zone_raw = self._win.evaluate_js("window.__hoverZone || null") if self._win else None
            zone = str(zone_raw).strip("'\"") if zone_raw else None
            if zone_raw is not None:
                self._emit("window.__hoverZone = null")   # 消费掉，避免残留影响下一次拖放
            dt = e.get("dataTransfer") or {}
            files = dt.get("files") or []
            self._drop_log("drop zone=%r files=%r" % (zone, [f.get("name") for f in files]))
            if zone == "file":
                self._emit("window.flagDrop()")   # 后端已接收，阻止前端超时兜底重复提示
                self._handle_drop_file(dt)
            elif zone == "folder":
                self._emit("window.flagDrop()")
                self._handle_drop_folder(dt)
            else:
                self._drop_log("drop 未命中任何拖放区（zone=%r）" % (zone_raw,))
        except Exception as ex:
            import traceback
            self._drop_log("drop 处理异常: %s\n%s" % (ex, traceback.format_exc()))

    def _drop_paths(self, data_transfer):
        """从事件取真实路径列表（pywebview 注入 pywebviewFullPath）。"""
        paths = []
        for f in data_transfer.get("files") or []:
            p = f.get("pywebviewFullPath") or f.get("path")
            if p:
                paths.append(p)
        return paths

    def _handle_drop_file(self, dt):
        paths = self._drop_paths(dt)
        if not paths:
            self._drop_log("file 区：未取到路径（files=%r）" % ([f.get("name") for f in dt.get("files") or []]))
            self._emit('toast("未能获取文件路径，请用「浏览…」选择文件")')
            return
        p = paths[0]
        self._drop_log("file 区路径: %s" % p)
        if not os.path.isfile(p):
            self._emit('toast("拖入的不是有效文件：%s")' % json.dumps(p, ensure_ascii=False))
            return
        r = self.load_path(p)
        if r.get("error"):
            self._emit("toast(%s)" % json.dumps(r["error"], ensure_ascii=False))
            return
        self._emit("applyFileFromBackend(%s)" % json.dumps(r, ensure_ascii=False))

    def _handle_drop_folder(self, dt):
        paths = self._drop_paths(dt)
        if not paths:
            self._drop_log("folder 区：未取到路径（files=%r）" % ([f.get("name") for f in dt.get("files") or []]))
            self._emit('toast("未能获取文件夹路径，请用「浏览…」选择文件夹")')
            return
        p = paths[0]
        self._drop_log("folder 区路径: %s" % p)
        if not os.path.isdir(p):
            self._emit("toast(%s)" % json.dumps("拖入的不是文件夹：" + p, ensure_ascii=False))
            return
        save_settings(p)
        self._emit("setOutDirFromBackend(%s)" % json.dumps(p, ensure_ascii=False))
        self._emit("toast(%s)" % json.dumps("输出目录已设为 " + p, ensure_ascii=False))

    def open_dir(self, d):
        if d and os.path.isdir(d):
            os.startfile(d)
        else:
            self._emit('toast("输出目录还不存在，拆分完成后会自动创建。")')

    def split(self, path, sheet, header_row, cols, out_dir):
        if self._splitting:
            return
        self._splitting = True

        def worker():
            q = queue.Queue()

            def progress_cb(i, t, k):
                q.put(("prog", i, t, k))

            def log_cb(m):
                q.put(("log", m))

            def drain():
                while True:
                    try:
                        item = q.get_nowait()
                    except queue.Empty:
                        break
                    if item[0] == "prog":
                        _, i, t, k = item
                        self._emit("onProgress(%d, %d, %s)" % (i, t, json.dumps(k, ensure_ascii=False)))
                    else:
                        self._emit("onLog(%s)" % json.dumps(item[1], ensure_ascii=False))

            try:
                created, errors, out = core.split_workbook(
                    path, sheet=sheet or None,
                    col_indices=[int(c) for c in cols],
                    header_row=int(header_row),
                    out_dir=out_dir or None,
                    progress_cb=progress_cb, log_cb=log_cb)
                drain()
                payload = {"created": created, "errors": errors, "out_dir": out}
                save_settings(out)          # 记忆输出目录
                self._emit("onDone(%s)" % json.dumps(payload, ensure_ascii=False))
            except Exception as e:
                drain()
                self._emit("onError(%s)" % json.dumps(str(e), ensure_ascii=False))
            finally:
                self._splitting = False

        threading.Thread(target=worker, daemon=True).start()


def _set_mica_hwnd(hwnd):
    """对指定句柄开启 Mica。成功返回 True。"""
    try:
        dwm = ctypes.windll.dwmapi
        val = ctypes.c_int(_DWMSBT_MAINWINDOW)
        if dwm.DwmSetWindowAttribute(hwnd, _DWMWA_SYSTEMBACKDROP_TYPE,
                                     ctypes.byref(val), ctypes.sizeof(val)) == 0:
            return True
        val = ctypes.c_int(1)
        if dwm.DwmSetWindowAttribute(hwnd, _DWMWA_MICA_EFFECT,
                                     ctypes.byref(val), ctypes.sizeof(val)) == 0:
            return True
    except Exception:
        pass
    return False


def run_gui():
    """GUI 入口（由 excel_splitter.py 无参数调用，或本文件直接运行）。"""
    api = Api()
    window = webview.create_window(
        WINDOW_TITLE,
        os.path.join(GUI_DIR, "index.html"),
        js_api=api,
        width=1060, height=720, min_size=(900, 620),
        background_color="#e8ecf1",
        transparent=True,   # WebView2 背景透明，真 Mica 才能透出；CSS 渐变负责兜底
    )
    api._win = window

    def on_loaded():
        # 页面加载完成后再绑定原生 DOM 拖放。早绑会挂到 about:blank 的旧 document 上，
        # 导航后监听器丢失，导致 WebView2 默认兜底行为（文件→下载 / 文件夹→开资源管理器）。
        try:
            from webview.dom import DOMEventHandler
            window.dom.document.events.dragover += DOMEventHandler(
                api._on_dom_dragover, prevent_default=True, debounce=200)
            window.dom.document.events.drop += DOMEventHandler(
                api._on_dom_drop, prevent_default=True, stop_propagation=True)
            api._dom_ok = True
        except Exception as e:
            api._dom_ok = False
            print("DOM 拖放绑定失败（将仅支持「浏览…」）: " + str(e))
        try:
            hwnd = ctypes.windll.user32.FindWindowW(None, WINDOW_TITLE)
            api._mica_ok = bool(hwnd) and _set_mica_hwnd(hwnd)
        except Exception:
            api._mica_ok = False
        window.evaluate_js("window.onReady(%s, %s)" %
                           ("true" if api._dom_ok else "false",
                            "true" if api._mica_ok else "false"))

    window.events.loaded += on_loaded
    webview.start(debug=False)


if __name__ == "__main__":
    run_gui()
