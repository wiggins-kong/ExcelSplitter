# Changelog

本项目的所有重要变更都记录在此文件中。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [v2.0.1] - 2026-09-05

### 新增

- **全新项目 Logo 与应用图标**：
  - 设计方向：钢蓝 / 雾蓝为主色、少量雾紫点缀，低饱和；图形用「一整张表沿列裂开成多张」的切割隐喻，与「按列拆分」的工具定位呼应
  - 定稿形态：**白色 squircle 圆角图标**（连续曲率圆角，Win11 / macOS 应用图标轮廓），主体居中、留白均匀，无文字
  - 新增 `assets/icons/`，收录设计源图（2048×2048）：白底版 / 深底版 / 纯图形透明底 / 首轮 4 个候选稿
  - 新增 `assets/icons/ExcelSplitter.ico`：16/24/32/48/64/128/256 多尺寸，PNG 压缩帧，圆角外真透明（深色任务栏下不露白方块）
  - 新增 `assets/icons/make_ico.py`：一键从设计源图重新生成 `.ico` 的脚本（需要 Pillow）
  - README 顶部展示 Logo，新增「图标资源」一节说明文件用途与重新生成方法

### 打包

- 构建命令新增 `--icon assets/icons/ExcelSplitter.ico`，exe / 任务栏 / 窗口图标统一为新版 Logo
- `ExcelSplitter.spec` / `ExcelSplitterDbg.spec` 的 EXE 段同步写入 `icon=`
- `.github/workflows/build-release.yml` 云端构建同步带 `--icon`，推 tag 产出的 Release exe 自动带图标

> 说明：ico 由「白色版 JPG 反解 alpha」生成——该源图是铺满画布的白色 squircle，圆角外被 JPG 压成黑色，据此用非黑区定底板、边界过渡带亮度做软 alpha、白区提纯白，得到带真透明的圆角图标。

## [v2.0.0] - 2026-09-03

### 重大变更：GUI 全面重构（Win11 Fluent）

- **技术栈更换**：tkinter / sv_ttk / tkinterdnd2 → **pywebview + WebView2**（Win10/11 自带运行时）
- 全新界面（Win11 Fluent 设计）：
  - **真 Mica 云母背景**：Win11 22H2+ 通过 `DwmSetWindowAttribute` 开启系统云母，WebView2 背景透明；旧系统自动回退 CSS 渐变模拟
  - 毛玻璃（Acrylic）卡片：`backdrop-filter` 半透明模糊材质
  - 微软雅黑字体、系统原生标题栏（Mica 透过标题栏）
- **布局重构**：双栏结构——左栏「① 选择文件 → ② 拆分设置 → ③ 输出位置」，右栏「④ 开始拆分（大号百分比 + 进度条）→ ⑤ 运行日志」
- 列选择为 HTML 表格，支持单击 / Ctrl / Shift 多选，实时「预计生成 N 个文件」
- 文件拖入窗口载入（HTML5 拖放，替代 tkdnd）
- 拆分后台线程执行，进度 / 日志实时推送到界面
- 输出目录记忆（`%APPDATA%\ExcelSplitter\settings.json`，沿用）
- 核心拆分逻辑 `excel_splitter.py` 与 CLI 用法**零改动**

### 新增

- **深色主题，随 Windows 深浅自动切换**：
  - 双套 CSS Design Tokens（`:root` 浅色 / `html[data-theme="dark"]` 深色），所有组件颜色走变量，切换零样式重复
  - 前端 `matchMedia('(prefers-color-scheme: dark)')` 实时跟随系统切换 + 后端注册表（`AppsUseLightTheme`）权威注入兜底
  - 标题栏 / 窗口边框随系统深浅：`DwmSetWindowAttribute(20, ImmersiveDarkMode)`
  - 深色系统下窗口采用纯色底 + CSS 深色渐变，规避深色 Mica 溢出浅底导致的内容区发白
  - 交互增强：点击列名即「勾选/取消勾选」（默认多选，不再清空其他列），Shift 仍支持范围选择
- 新增 3 个选列工具按钮：**全选 / 反选 / 取消选择**（与「工作表 / 表头行」同排、靠右对齐，载入文件后才可用）
- **「选择文件」模块新增「清除」按钮**：一键清空文件选择框 + 拆分设置的列列表 / 工作表 / 表头行（预计生成数隐藏），回到初始状态
- **「运行日志」模块新增「清空日志」按钮**：清空日志区（日志为空时自动禁用）

### 修复

- **文件夹拖到「输出位置」失败（"未能获取文件夹路径"）**：
  - 根因：前端拖放区 `drop` handler 调用了 `e.stopPropagation()` 并提前把 `window.__hoverZone` 置空，导致 drop 事件**不冒泡到 document**，pywebview 原生 DOM 监听（`window.dom.document.events.drop`）永远收不到事件，`pywebviewFullPath` 真实路径注入链路（`FilesDropped` → 文件名匹配）根本未执行；前端兜底对文件夹又拿不到 `File.path`（WebView2 沙箱恒为 null），只能弹提示引导「浏览…」
  - 修复：前端拖放区 `drop` handler 不再 `stopPropagation` / 不清空 `__hoverZone`，让事件冒泡到 document 交给后端原生监听路由；后端 `_on_dom_drop`「消费」式读取 `__hoverZone` 后立即清空，并在命中拖放区时先置 `__dropFeedback` 阻止前端超时兜底重复提示
  - 顺带修复：文件夹拖放成功的 toast 此前因 JS 嵌套引号是语法错误（`toast("已设为 "C:\...""）`）而不显示，已统一为 `toast(json.dumps(...))` 写法
  - 效果：文件夹拖入「输出位置」经原生链路取到真实路径并设为输出目录（记忆）；文件拖放链路同样受益（不再等 1.2s 超时兜底）

### 打包

- 依赖变更为：`openpyxl xlrd pyinstaller pywebview`（不再需要 tkinter / sv-ttk / tkinterdnd2）
- 构建命令新增 `--collect-all webview --add-data "webgui;webgui"`
- 运行时依赖 Windows WebView2 Runtime（Win10/11 系统自带）

### 构建与分发（CI / 双托管）

- **Release 附件文件名带版本号**：构建产物改为 `ExcelSplitter-<tag>.exe`（如 `ExcelSplitter-v2.0.0.exe`）
- **新增 Gitee 源码镜像**：`.github/workflows/sync-to-gitee.yml` 在 push 时自动把全部分支与 tags 同步到 Gitee（`gitee.com/wiggins-kong/ExcelSplitter`），全程云端、无需本地命令
- **Gitee Release 附件改为手动同步**：曾尝试云端自动上传（推 tag → Gitee OpenAPI 建 Release → 下载 GitHub 附件上传 Gitee），因构建机跨境（美国 → 境内）上传大文件不稳定而放弃：Node fetch 默认 5 分钟 body 超时，换 curl 重试 3 次后 Gitee 仍报 `file is missing`（服务端未收全文件体）。Gitee 附件改手动：GitHub Releases 下载 exe → Gitee「发行版 · 编辑发布」上传

## [v1.1.0] - 2026-09-02

### 界面重构（全新 GUI）

- 引入 **sv_ttk 2.6.1**（Sun Valley 主题），控件全面现代化：圆角、Windows 11 原生观感
- 新布局：深蓝 header + 徽章式步骤卡片（① 选择文件 / ② 拆分设置 / ③ 输出位置 / ④ 运行日志）
- 列选择由 Listbox 升级为 **Treeview 表格**（列号 / 列名 / 示例值，支持多选、垂直滚动）
- 底部操作区：状态指示灯 + 进度条 + 大号 Accent 按钮；运行日志升级为时间+消息表格
- 窗口默认 824×700，可自由拉伸自适应
- 新增**文件拖拽**支持（tkinterdnd2 2.6.2 / tkdnd 2.10.1）：文件直接拖到窗口即载入
- 新增**预计生成数**：选中拆分列后实时估算将生成的文件数
- 输出目录**记忆**：上次的输出目录自动记住（`%APPDATA%\ExcelSplitter\settings.json`）
- 拆分任务改为**后台线程执行**，界面不再卡顿；完成提示「成功生成 N 个文件」+「打开输出目录」按钮

### 重构（结构）

- 拆分为两个文件：`excel_splitter.py`（核心逻辑 + CLI，无 GUI 依赖）+ `excel_splitter_gui.py`（GUI）
- CLI 入口与用法不变；GUI 需系统 Python（tkinter）+ sv_ttk + tkinterdnd2

### 修复

- **Tk 9.0 + sv_ttk 背景重置**：Tk 9.0 会在布局刷新后用主题默认色覆盖 tk 部件显式背景，导致自绘 header 变灰；全部自定义颜色改走 `ttk.Style`
- 高 DPI：进程 DPI 感知改为 PerMonitorV2（`SetProcessDpiAwarenessContext`）三级回退，配合既有 `dpi_manifest.xml`

## [v1.0.1] - 2026-09-02

### 新增

- 支持 `.xls` 旧格式输入（引入 `xlrd`，按扩展名自动分派读取引擎）：
  - `.xls` → xlrd；`.xlsx` / `.xlsm` → openpyxl（保持不变）
  - `.xls` 日期单元格自动转为 datetime，与 `.xlsx` 读取行为一致
  - GUI 文件选择过滤器、CLI 帮助文本同步支持 `.xls`
- 新增统一读取层 `list_sheets()` / `read_sheet_rows()`，集中管理格式分派与读取逻辑

### 修复

- **修复表头只读到第一个单元格的问题**：部分工具导出的 `.xlsx` 内部 `<dimension>` 声明错误（如实际 37 列却声明 `A1`），openpyxl `read_only=True` 模式信任该声明会把工作表截断成 1×1，导致列列表只显示第一列表头（如「订单编号」）。已废弃 `read_only` 模式，改用常规模式按实际单元格解析
- **修复打包后 exe 缺失 xlrd**：构建命令新增 `--clean --hidden-import xlrd` 强制包含；同步修复 GitHub Actions workflow（此前依赖安装缺 xlrd、打包命令缺 hidden-import，Release exe 分发后无法读取 `.xls`）

## [v1.0.0] - 2026-08-20

初始发布：WPS 表格 JS 宏 `SplitByNonAdjacentSelection` 的独立桌面版实现。

- 按多列组合分组拆分 Excel（GUI + CLI 双模式）
- 强力数据清洗：空值 → 空白、去首尾空格、去单元格内换行 / 制表符
- 非破坏性：原表只读不改；拆分文件保留完整表头与所有列
- 非法文件名处理、自适应列宽、高 DPI 感知
- GitHub Actions 推 `v*` tag 自动打包发布 exe