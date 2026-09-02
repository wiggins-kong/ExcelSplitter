# DEVELOPMENT.md — 开发记录 / 接续指南

> 用途：开发者**换电脑 / 新开会话**时快速接上项目状态。
> 文档分工：`README.md` = 用户使用手册；`CHANGELOG.md` = 版本变更；`DEVELOPMENT.md` = 开发进度与踩坑记录。

## 项目一句话

ExcelSplitter：按指定列把 Excel 总表拆分成多个独立 `.xlsx` 的桌面小工具（tkinter GUI + CLI，openpyxl / xlrd 读取，PyInstaller 单文件分发，WPS JS 宏的独立实现）。

## 当前状态（2026-09-02）

- **版本**：`v1.0.1` 已打 tag（指向提交 `7e95fdf`），**尚未推送到 GitHub**（用户自行推送）
- **功能**：支持 `.xlsx` / `.xlsm` / `.xls` 输入，按多列组合拆分输出 `.xlsx`
- **已验证**：CLI 拆分 `.xls` / `.xlsx` 均实测通过（98 地区拆分、日期转换、表头完整性）
- **未验证**：GUI 模式尚未用真实带 tkinter 的 Python 重打包实测

## 必知坑（踩过，请别再踩）

1. **openpyxl 禁用 `read_only=True`**：部分工具导出的 xlsx 内部 `<dimension>` 声明错误（真实 37 列却声明 `A1`），read_only 模式信任该声明会把表截成 1×1，表现为"表头只读到第一个单元格"。统一走 `read_sheet_rows()` 常规模式。
2. **xlrd 为可选导入，PyInstaller 不保证自动打包**：打包机没装 xlrd 或命令缺 `--hidden-import`，exe 分发后报「读取 .xls 文件需要 xlrd 库」。打包必须：打包机 `pip install xlrd` + 命令带 `--clean --hidden-import xlrd`。
3. **GitHub Actions workflow 与本地打包命令需同步维护**：`.github/workflows/build-release.yml` 里依赖安装和打包参数都要有 xlrd（已修，勿回退）。
4. **GUI 打包必须用带 tkinter 的 Python**：本机 managed 隔离 venv（Python 3.13）无 tkinter，打出的 exe 只能跑 CLI；GUI 版需用系统自带 Python 打包。

## 关键代码位置

| 位置 | 说明 |
| --- | --- |
| `read_sheet_rows()` / `list_sheets()` | 统一读取层，按扩展名分派：`.xls` → xlrd，`.xlsx/.xlsm` → openpyxl |
| `split_workbook()` | 拆分主逻辑（元组作分组键，避免分隔符与数据冲突） |
| `_require_xlrd()` | xlrd 缺失时的友好报错 |
| `.github/workflows/build-release.yml` | push `v*` tag 触发自动打包 + 创建 Release |

## 常用命令

```bash
# ① CLI 拆分
python excel_splitter.py 文件.xlsx --cols 1,3 --out 输出目录 --header 1

# ② 构建 exe（必须含 --clean --hidden-import xlrd）
pip install openpyxl xlrd pyinstaller
pyinstaller --onefile --windowed --name ExcelSplitter --noupx \
  --manifest dpi_manifest.xml --clean --hidden-import xlrd excel_splitter.py

# ③ 打包后立即自检
dist\ExcelSplitter.exe "测试.xls" --cols 1 --out 自检目录

# ④ 发新版（推 tag 即触发 Actions Release）
git add -A && git commit -m "feat: ..."
git tag vX.Y.Z
git push origin main && git push origin vX.Y.Z
```

## 发布流程提醒

- 推送 `v*` tag → Actions 自动构建并把 `ExcelSplitter.exe` 挂到 Release
- 发版前需同步：`CHANGELOG.md` 补发布日期、本文件「当前状态」更新
- `v1.0.1` 待用户推送（`git push origin main` + `git push origin v1.0.1`）

## 项目约定 / 用户偏好

- 中文沟通、极简指令、**直接执行少反问**
- 不喜欢冗余文件与代码（`build.bat` 曾按用户要求删除）；保持文档精简
- 仓库不收录 `dist/` 产物与 `.workbuddy/`（均已 .gitignore）
- git 身份：本仓库 local 配置为 `WorkBuddy User <user@workbuddy.local>`（与历史提交一致）