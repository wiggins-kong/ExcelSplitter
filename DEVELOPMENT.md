# DEVELOPMENT.md — 开发记录 / 接续指南

> 用途：开发者**换电脑 / 新开会话**时快速接上项目状态。
> 文档分工：`README.md` = 用户使用手册；`CHANGELOG.md` = 版本变更；`DEVELOPMENT.md` = 开发进度与踩坑记录。

## 项目一句话

ExcelSplitter：按指定列把 Excel 总表拆分成多个独立 `.xlsx` 的桌面小工具（tkinter GUI + CLI，openpyxl / xlrd 读取，PyInstaller 单文件分发，WPS JS 宏的独立实现）。

## 当前状态（2026-09-06）

- **版本**：`v2.1.0`（2026-09-06）：**兼容扩展名造假的 ".xls" + 表头不透明**。① 读取分派从「按扩展名」改为「嗅探文件头」（`excel_splitter.py: _sniff_format()`）：`PK\x03\x04`→xlsx 走 openpyxl（扩展名不符时 `_load_xlsx()` 走 BytesIO 字节流绕过 openpyxl 扩展名校验）、`\xd0\xcf\x11\xe0`→旧版 xls 走 xlrd，修复伪装 .xls 报「Excel xlsx file; not supported」；② 选列表格表头 `--thead-bg` 深浅两套均改不透明实色（`#222426` / `#f3f6fa`），修复滚动时行文字与表头重叠。已验证：伪装 .xls CLI 拆分 11 文件 0 失败、GUI 冒烟通过；本地已重打包 `dist/ExcelSplitter.exe`。**待发布**：推 tag 触发 Actions Release
- **上一版**：`v2.0.1`（2026-09-05）：**全新项目 Logo + 应用图标**。白色 squircle 圆角图标定稿，新增 `assets/icons/`（设计源图 + `ExcelSplitter.ico` + `make_ico.py` 生成脚本），打包命令 / 两个 spec / CI workflow 全部带上 `--icon`。已发布 Release（附件 `ExcelSplitter-v2.0.1.exe`）
- **上上版**：`v2.0.0`（2026-09-03）：pywebview GUI 全面重构 + 拖放/选列交互增强 + 深色主题随系统切换；已发布 Release（附件 `ExcelSplitter-v2.0.0.exe`）
- **代码结构**：`excel_splitter.py`（核心 + CLI，零改动）+ `excel_splitter_gui_web.py`（pywebview 启动器 + JS API 桥）+ `webgui/index.html`（界面）
- **GUI 特性**：真 Mica（DwmSetWindowAttribute，Win11 22H2+；失败回退 CSS 渐变）、Acrylic 毛玻璃卡片、双栏布局（左配置/右执行与日志）、HTML 表格多选（默认点击多选 / Shift 范围选 / 全选·反选·取消选择）、HTML5 拖放（文件/文件夹真实路径）、预计生成数、目录记忆、后台线程拆分、清除文件、清空日志、**深浅双主题随 Windows 系统切换**
- **主题机制**：`webgui/index.html` 双套 CSS Tokens（`:root` 浅色 / `html[data-theme="dark"]` 深色，所有组件颜色全走变量）；前端 `matchMedia('(prefers-color-scheme: dark)')` 实时切换 + 后端注册表权威注入（onReady 三参 + `get_theme()` 兜底复查）；`transparent=not dark`——深色纯色底 + CSS 渐变，浅色透明 Mica；`DwmSetWindowAttribute(20, ImmersiveDarkMode)` 切标题栏/边框。深色初版见 `design/dark_theme_demo.html`
- **已验证**：源码 GUI 启动冒烟（含深色改动，09-03 00:0x 重启存活无 traceback）；源码/打包后 JS、Python 语法检查全过；exe 重打包 2 次（00:13/00:25）均完成；exe CLI 拆分 3 文件 0 失败；深色根因已修（见必知坑 9/10）；**用户真机确认深色观感正常（09-03 00:30）**；已 commit + 本地 tag `v2.0.0`（未推送）
- **未验证**：GUI 全流程人工操作（载入→选列→估算→拆分）；真机拖拽（含文件夹放行）
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
9. **深色系统下别用透明 Mica（v2.0.1 深色主题踩穿）**：部分环境深色 Mica 溢出的底色是浅灰白，`transparent=True + 内容透明` 会把内容区整片染白（标题栏深、内容白就是它）。对策：`create_window(transparent=not dark, background_color="#0f1113" if dark else "#e8ecf1")`，深色走 CSS 渐变兜底、浅色走 Mica；`on_loaded` 里 `mica_ok = bool(hwnd) and (not dark) and _set_mica_hwnd(hwnd)`。
10. **深色主题判定双保险**：前端 `matchMedia('(prefers-color-scheme: dark)')` 在 WebView2 里可能滞后/失准 → `onReady` 第三参由后端读注册表 `AppsUseLightTheme` 权威注入；再加 `Api.get_theme()` 让前端 `applyBackendTheme()` 主动复查（防 onReady 时序不达白屏）。诊断日志写 `%TEMP%\ExcelSplitter\theme.log`（`_theme_log()`）。
11. **Gitee 双托管**：源码自动同步走 `.github/workflows/sync-to-gitee.yml`（push 触发，`git push gitee --all --tags`，认证 URL `https://oauth2:${GITEE_TOKEN}@gitee.com/wiggins-kong/ExcelSplitter.git`，Secret `GITEE_TOKEN`，Gitee 私人令牌须勾 projects）。**Release 附件云端自动同步已放弃**：云端 runner（美国 azure）跨境传 15MB multipart 到 Gitee 多次尝试均失败（Node fetch 默认 body 超时 300s；换 curl `--retry 3` 后 Gitee 仍报 `{"messages":["file is missing"]}`= 服务端没收全文件体），非代码问题。**现行方案（v2.1.0 起）：本地一键脚本 `scripts/publish_gitee_release.py`**——发版后在本机跑 `python scripts/publish_gitee_release.py v2.1.0 [exe路径]`，自动复制 GitHub Release 描述 + 建 Gitee Release + 直连上传 exe（本地网络无跨境问题）；需环境变量 `GITEE_TOKEN`；exe 省略时先找 `dist/` 再从 GitHub Release 下载（用 `GITHUB_TOKEN` 或 git 凭证管理器）。踩过的 Gitee OpenAPI 坑：**access_token 必须放 query/formData，放 JSON body 返回 401**；**创建 Release 必须显式传 `target_commitish`**；**Gitee API 无删除附件接口**，同名附件已存在时脚本会提示去网页手动删旧的重跑。

12. **设计稿的「透明底 PNG」未必有底板**：美图设计室返回的 `icon_transparent.png`（自称圆角图标·透明底）实际只有图形、没有圆角底板，直接转 ico 会丢外框。而 `icon-white-bg.jpg` 是铺满画布的白色 squircle、圆角外被 JPG 压成**黑色**，反而能反解 alpha（非黑区定底板 + 膨胀/腐蚀差集取边界带用亮度做软 alpha + 白区提纯白）。换图前先验像素：`corner(5,5)` 与 `center` 的 alpha/亮度对比一下就知道有没有底板。详见 `assets/icons/make_ico.py`。
13. **本地打包要带上 `.workbuddy/gui_demo/deps`**：`pywebview` / `Pillow` 装在项目内的 `deps` 目录（系统 Python 3.14 全局只有 openpyxl / xlrd / pyinstaller），打包命令前加 `PYTHONPATH=.workbuddy/gui_demo/deps`，否则 `ModuleNotFoundError: webview`。（注：本机后来已全局装齐 pywebview，此坑仅旧机器适用）
14. **".xls" 扩展名不可信，读取分派要嗅探文件头（v2.0.2 踩穿）**：不少系统导出的 ".xls" 实为 xlsx（ZIP 容器，头 `PK\x03\x04`），按扩展名丢给 xlrd 会报「Excel xlsx file; not supported」。对策：`excel_splitter.py: _sniff_format()` 按头 8 字节分派（`PK\x03\x04`→openpyxl、`\xd0\xcf\x11\xe0`→xlrd，嗅探不出回退扩展名）；且 openpyxl 会按扩展名校验格式，扩展名不符时必须走 `_load_xlsx()` 字节流（BytesIO）加载。`list_sheets` / `read_sheet_rows` 已全部切换，GUI 与 CLI 自动受益。另：`--windowed` exe 里 CLI 抛未捕获异常会弹错误对话框挂住进程（不会自己退出），冒烟脚本要防卡死。

## 关键代码位置

| 位置 | 说明 |
| --- | --- |
| `excel_splitter.py: read_sheet_rows()` / `list_sheets()` | 统一读取层，按扩展名分派：`.xls` → xlrd，`.xlsx/.xlsm` → openpyxl |
| `excel_splitter.py: split_workbook()` | 拆分主逻辑（元组作分组键，避免分隔符与数据冲突） |
| `excel_splitter.py: group_keys()` | 只统计分组不写文件，供 GUI 实时显示预计生成数 |
| `excel_splitter_gui_web.py: Api` | pywebview JS API 桥（pick_file / load_path / get_columns / estimate / split 等） |
| `excel_splitter_gui_web.py: run_gui()` | 窗口创建（transparent=True）+ Mica DWM 调用 |
| `webgui/index.html` | 整个界面（HTML/CSS/JS，与 Python 通过 pywebview.api / window.onXxx 事件通信） |
| `.github/workflows/build-release.yml` | push `v*` tag 触发自动打包 + 创建 GitHub Release（附件名带版本号 `ExcelSplitter-<tag>.exe`） |
| `.github/workflows/sync-to-gitee.yml` | push 触发，自动同步全部分支 + tags 到 Gitee（源码镜像，云端执行） |

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

# ⑤ Gitee Release 镜像（等 Actions 构建完成后本地跑；需环境变量 GITEE_TOKEN）
python scripts/publish_gitee_release.py vX.Y.Z
```

## 发布流程提醒

- 推送 `v*` tag → Actions 自动构建并把 `ExcelSplitter-<tag>.exe` 挂到 GitHub Release（Release notes 自动取自 CHANGELOG 对应版本段）；源码随即自动同步到 Gitee（sync-to-gitee）
- **Gitee Release 用本地脚本镜像**：`python scripts/publish_gitee_release.py vX.Y.Z`（详建坑 11）；云端自动传附件已验证不可行，别再试
- 发版前需同步：`CHANGELOG.md` 补发布日期、本文件「当前状态」更新
- `v2.1.0`：GitHub Release 已发布（notes 自动提取生效）；Gitee Release 待用脚本补传 exe

## 项目约定 / 用户偏好

- 中文沟通、极简指令
- 不喜欢冗余文件与代码（`build.bat` 曾按用户要求删除）；保持文档精简
- 仓库不收录 `dist/` 产物与 `.workbuddy/`（均已 .gitignore）
- git 身份：本仓库 local 配置为 `WorkBuddy User <user@workbuddy.local>`（与历史提交一致）
- 双托管：GitHub 主仓（`github.com/wiggins-kong/ExcelSplitter`）+ Gitee 镜像（`gitee.com/wiggins-kong/ExcelSplitter`）；源码自动同步、Gitee 附件手动；Gitee Token 存 GitHub Actions Secret `GITEE_TOKEN`（私人令牌勾 projects）