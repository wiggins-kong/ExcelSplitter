# DEVELOPMENT.md — 开发记录 / 接续指南

> 用途：开发者**换电脑 / 新开会话**时快速接上项目状态。
> 文档分工：`README.md` = 用户使用手册；`CHANGELOG.md` = 版本变更；`DEVELOPMENT.md` = 开发进度与踩坑记录。

## 项目一句话

ExcelSplitter：按指定列把 Excel 总表拆分成多个独立 `.xlsx` 的桌面小工具（tkinter GUI + CLI，openpyxl / xlrd 读取，PyInstaller 单文件分发，WPS JS 宏的独立实现）。

## 当前状态（2026-09-02）

- **版本**：`v1.1.0` 界面重构已本地完成并全流程自测通过（GUI + CLI + exe），**未打 tag / 未推送**（用户试用满意后再推）
- **代码结构**：已拆两文件——`excel_splitter.py`（核心 + CLI，无 GUI 依赖）+ `excel_splitter_gui.py`（sv_ttk 新界面）
- **GUI 新特性**：sv_ttk 2.6.1 圆角主题、深蓝 header、徽章步骤卡、Treeview 列选择（列名+示例值）、文件拖拽（tkinterdnd2 0.6.2 / tkdnd 2.10.1）、预计生成数实时估算、输出目录记忆（%APPDATA%\ExcelSplitter\settings.json）、拆分后台线程不卡 UI
- **已验证**：源码 GUI 自动化全流程（载入→选列→估算→拆分→6 文件 0 失败）；exe CLI 拆分 6 文件；exe GUI 窗口启动 + sv_ttk 渲染正常
- **未验证**：真机拖拽（源码/打包均验证 tkdnd 加载成功，尚未人工拖文件实测）

## 必知坑（踩过，请别再踩）

1. **openpyxl 禁用 `read_only=True`**：部分工具导出的 xlsx 内部 `<dimension>` 声明错误（真实 37 列却声明 `A1`），read_only 模式信任该声明会把表截成 1×1，表现为"表头只读到第一个单元格"。统一走 `read_sheet_rows()` 常规模式。
2. **xlrd 为可选导入，PyInstaller 不保证自动打包**：打包机没装 xlrd 或命令缺 `--hidden-import`，exe 分发后报「读取 .xls 文件需要 xlrd 库」。打包必须：打包机 `pip install xlrd` + 命令带 `--clean --hidden-import xlrd`。
3. **GitHub Actions workflow 与本地打包命令需同步维护**：`.github/workflows/build-release.yml` 依赖安装和打包参数要含全部五项（openpyxl/xlrd/pyinstaller/sv-ttk/tkinterdnd2），且带 `--hidden-import xlrd --hidden-import sv_ttk --collect-all sv_ttk --collect-all tkinterdnd2`。
4. **GUI 打包必须用带 tkinter 的 Python**：本机 managed 隔离 venv（Python 3.13）无 tkinter，打出的 exe 只能跑 CLI；GUI 版需用 `AppData\Local\Programs\Python\Python314`（Tk 9.0）打包。
5. **`sv_ttk` 必须与 tkinterdnd2 同目录才被 PyInstaller 找到**：sv_ttk 若是手工解压放置，`--paths` 必须指向它所在目录；统一放 `deps/` 并 `--hidden-import sv_ttk --collect-all sv_ttk`（缺 `sv.tcl`/`theme` 数据会运行时报 No module named 'sv_ttk'）。
6. **tkinterdnd2 打包必须 `--collect-all tkinterdnd2`**：tkdnd 是数据（Tcl）+ 二进制（dll），漏收则拖拽静默失效；且**tkdnd 2.10.1 才兼容 Tk 9.0**（旧版报错）。
7. **Tk 9.0 + sv_ttk 会重置 tk 部件背景**：`update_idletasks()` 后 tk 部件（Frame/Label）显式 `-background` 被主题默认值覆盖（变 #fafafa）。**自定义颜色一律走 ttk.Style**（`style.configure("Header.TFrame", background=...)`），不用 tk 部件显式设色；Canvas 背景被重置但近白无感可容忍。
8. **Tk 9.0 PIXEL 选项 bug**：部分像素选项值不保留（官方已知），UI 尺寸用字符/相对单位，避免写死像素。

## 关键代码位置

| 位置 | 说明 |
| --- | --- |
| `excel_splitter.py: read_sheet_rows()` / `list_sheets()` | 统一读取层，按扩展名分派：`.xls` → xlrd，`.xlsx/.xlsm` → openpyxl |
| `excel_splitter.py: split_workbook()` | 拆分主逻辑（元组作分组键，避免分隔符与数据冲突） |
| `excel_splitter.py: group_keys()` | 只统计分组不写文件，供 GUI 实时显示预计生成数（口径与 split_workbook 一致） |
| `excel_splitter.py: _require_xlrd()` | xlrd 缺失时的友好报错 |
| `excel_splitter_gui.py: ExcelSplitterApp` | 整个新界面（sv_ttk 样式、四步卡片、拖拽、记忆、线程拆分） |
| `.github/workflows/build-release.yml` | push `v*` tag 触发自动打包 + 创建 Release（依赖与打包参数已同步 v1.1.0） |

## 常用命令

```bash
# ① CLI 拆分
python excel_splitter.py 文件.xlsx --cols 1,3 --out 输出目录 --header 1

# ② 构建 exe（必须包含全部五项依赖与 collect 参数）
pip install openpyxl xlrd pyinstaller sv-ttk tkinterdnd2
pyinstaller --onefile --windowed --name ExcelSplitter --noupx \
  --manifest dpi_manifest.xml --clean \
  --hidden-import xlrd --hidden-import sv_ttk \
  --collect-all sv_ttk --collect-all tkinterdnd2 \
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
- `v1.1.0` 待用户本地试用满意后推送（`git tag v1.1.0` + push）

## 项目约定 / 用户偏好

- 中文沟通、极简指令、**直接执行少反问**
- 不喜欢冗余文件与代码（`build.bat` 曾按用户要求删除）；保持文档精简
- 仓库不收录 `dist/` 产物与 `.workbuddy/`（均已 .gitignore）
- git 身份：本仓库 local 配置为 `WorkBuddy User <user@workbuddy.local>`（与历史提交一致）