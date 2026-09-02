# DEVELOPMENT.md — 开发记录 / 接续指南

> 用途：开发者**换电脑 / 新开会话**时快速接上项目状态。
> 文档分工：`README.md` = 用户使用手册；`CHANGELOG.md` = 版本变更；`DEVELOPMENT.md` = 开发进度与踩坑记录。

## 项目一句话

ExcelSplitter：按指定列把 Excel 总表拆分成多个独立 `.xlsx` 的桌面小工具（tkinter GUI + CLI，openpyxl / xlrd 读取，PyInstaller 单文件分发，WPS JS 宏的独立实现）。

## 当前状态（2026-09-02）

- **版本**：`v2.0.1`（基于 v2.0.0 的 pywebview GUI）：修复文件夹拖放失败 + 选列交互增强（默认多选、全选/反选/取消选择、清除文件、清空日志），**未打 tag / 未推送**
- **代码结构**：`excel_splitter.py`（核心 + CLI，零改动）+ `excel_splitter_gui_web.py`（pywebview 启动器 + JS API 桥）+ `webgui/index.html`（界面）
- **GUI 特性**：真 Mica（DwmSetWindowAttribute，Win11 22H2+；失败回退 CSS 渐变）、Acrylic 毛玻璃卡片、双栏布局（左配置/右执行与日志）、HTML 表格多选（默认点击多选 / Shift 范围选 / 全选·反选·取消选择）、HTML5 拖放（文件/文件夹真实路径）、预计生成数、目录记忆、后台线程拆分、清除文件、清空日志
- **已验证**：源码 GUI 启动冒烟；exe CLI 拆分 2 文件 0 失败（23:45 重建）；exe GUI 启动存活；拖放修复经 pywebview 源码链路确认 + 单测/集成测试（路由、目录记忆、emit 合法性）
- **未验证**：GUI 全流程人工操作（载入→选列→估算→拆分）；真机拖拽（含文件夹放行）；Mica 视觉效果（需人工看）
- **Python 3.14 位置已变**：`AppData\Local\Python\pythoncore-3.14-64`（GUI 打包用，含 tkinter 无所谓了；依赖需单独装：openpyxl/xlrd/pyinstaller/pywebview）

## 必知坑（踩过，请别再踩）

1. **openpyxl 禁用 `read_only=True`**：部分工具导出的 xlsx 内部 `<dimension>` 声明错误（真实 37 列却声明 `A1`），read_only 模式信任该声明会把表截成 1×1。统一走 `read_sheet_rows()` 常规模式。
2. **xlrd 为可选导入，PyInstaller 不保证自动打包**：打包命令必须带 `--clean --hidden-import xlrd`。
3. **pywebview 打包必须 `--collect-all webview --add-data "webgui;webgui"`**：缺 collect-all 报 WebView2Loader 缺失；缺 add-data GUI 白屏（找不到 index.html）。
4. **WebView2 Runtime 是运行时依赖**：Win10/11 自带；极老系统需装 [WebView2 Runtime]。
5. **真 Mica 需 `transparent=True` + DWM 调用**：pywebview 的 `create_window(transparent=True)` 让 WebView2 背景透明，Mica 才能透出；`on_loaded` 里 FindWindowW 按标题拿 hwnd 再 `DwmSetWindowAttribute(38, DWMSBT_MAINWINDOW=2)`。失败时前端 `body.no-mica` 用 CSS 渐变兜底——**两者都会工作，视觉需人工确认**。
6. **GitHub Actions workflow 已同步 v2.0.0 依赖与打包参数**。
7. **CLI 自检命令**：`dist\ExcelSplitter.exe 测试.xlsx --cols 1,2 --out 输出目录`
8. **WebView2 拖放三座坑（已踩穿）**：
   - 页面 `File.path` 恒为 `null`（沙箱），要真实路径必须走 pywebview 原生 DOM 事件（`webview.dom` + `DOMEventHandler`，机制见关键代码位置）。成功链路：注入监听 → `_jsApiCallback('pywebviewEventHandler',{event})` → edgechromium `postMessageWithAdditionalObjects('FilesDropped')` 收集真实路径到 `webview.dom._dnd_state['paths']` → util.py 按文件名匹配注入 `pywebviewFullPath`
   - **DOM 监听必须等页面加载完成后绑**（`window.events.loaded`）：早绑挂到 about:blank 旧 document，导航后丢失 → WebView2 默认行为接管（文件→下载、文件夹→开资源管理器）
   - `evaluate_js` 返回值带引号（如 `'file'`），与字符串比较前先 `strip("'\"")`；异常不要静默吞，写 `%TEMP%\ExcelSplitter\drop.log` 诊断日志
   - **前端拖放区 drop handler 禁止 `stopPropagation()`、禁止提前清 `__hoverZone`**（v2.0.1 踩穿）：事件必须冒泡到 document 才能触发 pywebview 原生监听并注入 `pywebviewFullPath`；zone 由后端 `_on_dom_drop`「消费」式读取后清空。曾因 stopPropagation 导致文件夹拖放 100% 失败、只弹「未能获取文件夹路径」（前端兜底拿不到 WebView2 沙箱里的 File.path）
   - 文件夹拖入取决于 WebView2 是否放行目录（多数放行）；不放行时前端 fallback 弹提示引导「浏览…」
   - 前端加超时自愈：拖放 1.2s 无后端反馈 → 文件区自动 base64 载入（`Api.load_blob`）、文件夹区提示；后端命中拖放区后须先 `window.flagDrop()` 防重复提示

## 关键代码位置

| 位置 | 说明 |
| --- | --- |
| `excel_splitter.py: read_sheet_rows()` / `list_sheets()` | 统一读取层，按扩展名分派：`.xls` → xlrd，`.xlsx/.xlsm` → openpyxl |
| `excel_splitter.py: split_workbook()` | 拆分主逻辑（元组作分组键，避免分隔符与数据冲突） |
| `excel_splitter.py: group_keys()` | 只统计分组不写文件，供 GUI 实时显示预计生成数 |
| `excel_splitter_gui_web.py: Api` | pywebview JS API 桥（pick_file / load_path / get_columns / estimate / split 等） |
| `excel_splitter_gui_web.py: run_gui()` | 窗口创建（transparent=True）+ Mica DWM 调用 |
| `webgui/index.html` | 整个界面（HTML/CSS/JS，与 Python 通过 pywebview.api / window.onXxx 事件通信） |
| `.github/workflows/build-release.yml` | push `v*` tag 触发自动打包 + 创建 Release（已同步 v2.0.0） |

## 常用命令

```bash
# ① CLI 拆分
python excel_splitter.py 文件.xlsx --cols 1,3 --out 输出目录 --header 1

# ② 构建 exe（依赖：openpyxl xlrd pyinstaller pywebview）
pyinstaller --onefile --windowed --name ExcelSplitter --noupx \
  --manifest dpi_manifest.xml --clean \
  --hidden-import xlrd --collect-all webview --add-data "webgui;webgui" \
  excel_splitter.py

# ③ 打包后立即自检（CLI 拆分 + GUI 窗口启动）
dist\ExcelSplitter.exe "测试.xls" --cols 1 --out 自检目录

# ④ 发新版（推 tag 即触发 Actions Release）
git add -A && git commit -m "feat: ..."
git tag vX.Y.Z
git push origin main && git push origin vX.Y.Z
```

## 发布流程提醒

- 推送 `v*` tag → Actions 自动构建并把 `ExcelSplitter.exe` 挂到 Release
- 发版前需同步：`CHANGELOG.md` 补发布日期、本文件「当前状态」更新
- `v2.0.0` 待用户本地试用满意后推送（`git tag v2.0.0` + push）

## 项目约定 / 用户偏好

- 中文沟通、极简指令、**直接执行少反问**
- 不喜欢冗余文件与代码（`build.bat` 曾按用户要求删除）；保持文档精简
- 仓库不收录 `dist/` 产物与 `.workbuddy/`（均已 .gitignore）
- git 身份：本仓库 local 配置为 `WorkBuddy User <user@workbuddy.local>`（与历史提交一致）