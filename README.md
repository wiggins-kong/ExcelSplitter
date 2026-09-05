# Excel 按列拆分工具（ExcelSplitter）

<p align="center">
  <img src="assets/icons/icon-white-bg.jpg" width="128" alt="ExcelSplitter 图标" />
</p>

把一张 Excel 总表，按你**指定的若干列**自动拆分成多个独立的 `.xlsx` 文件。

支持输入 `.xlsx` / `.xlsm` / `.xls`，输出统一为 `.xlsx`。

这是 WPS 表格 JS 宏 `SplitByNonAdjacentSelection` 的**独立桌面版实现**——不依赖 WPS，双击即用。

## 功能特性

- **Win11 Fluent 界面**：pywebview + WebView2 实现，真 Mica 云母背景（Win11 22H2+，旧系统自动回退 CSS 模拟）、毛玻璃卡片、微软雅黑字体
- **深浅双主题，随 Windows 自动切换**：深色 / 浅色两套 Design Tokens + 标题栏同步，系统切主题程序即时跟随
- **双栏布局**：左栏「选择文件 → 拆分设置 → 输出位置」，右栏「开始拆分（进度）→ 运行日志」
- **多格式输入**：支持 `.xlsx` / `.xlsm` / `.xls`（.xls 由 xlrd 读取，日期单元格自动转 datetime）
- **按多列组合分组**：选中「地区」「产品」两列 → 自动拆成 `华东 - 苹果.xlsx`、`华北 - 香蕉.xlsx`……每个唯一组合生成一个文件
- **灵活选列**：点击列名即多选 / 反选（默认不打断已有选择），支持 Shift 范围选，配「全选 / 反选 / 取消选择」工具按钮
- **强力数据清洗**：空值 → `空白`、去首尾空格、去单元格内换行 / 制表符（专门解决「看着一样却没分到一组」的坑）
- **非破坏性**：原表只读不改，仅新建 / 保存拆分结果
- **完整保留**：每个拆分文件含完整表头与所有列
- **非法文件名处理**：`\ / * ? " < > |` → `-`、冒号替换、长度截断到 150
- **拖拽载入**：把 Excel 文件直接拖进窗口即载入；文件夹拖到「输出位置」即设为输出目录
- **预计生成数**：选中拆分列后实时显示将生成的文件数
- **记忆输出目录**：上次的输出目录自动记住
- **随手清整**：「清除」一键回到初始状态、「清空日志」随时清掉运行记录

## 使用方法

### 图形界面（推荐）

1. 双击 `dist/ExcelSplitter.exe`（需自行构建，见下方）
2. 选择文件：「浏览」或直接把文件**拖进窗口**（`.xlsx` / `.xlsm` / `.xls`）
3. 选择工作表、确认表头行（默认第 1 行）
4. 在「拆分设置」里用 **单击 / Ctrl / Shift** 多选拆分列，实时显示「预计生成 N 个文件」
5. 输出位置自动记忆上次目录；也可「浏览…」更换，「打开目录」直达结果
6. 点「开始拆分」→ 后台线程执行不卡界面，进度条与运行日志实时更新

### 命令行

```bash
ExcelSplitter.exe 文件.xlsx --cols 1,3 --out 输出目录 --header 1
```

| 参数 | 说明 |
| --- | --- |
| `input` | 输入 xlsx/xls 路径（位置参数） |
| `--cols` | 拆分列（1-based，逗号分隔，如 `1,3`） |
| `--sheet` | 指定工作表名 |
| `--header` | 表头行号，默认 `1` |
| `--out` | 输出目录 |

## 从源码构建

需要 **Windows 10/11 + Python 3.10+**（GUI 基于 pywebview + 系统 WebView2 运行时）。

```bash
pip install openpyxl xlrd pyinstaller pywebview
pyinstaller --onefile --windowed --name ExcelSplitter --noupx --manifest dpi_manifest.xml --clean \
  --icon assets/icons/ExcelSplitter.ico \
  --hidden-import xlrd --collect-all webview --add-data "webgui;webgui" \
  excel_splitter.py
```

- `--collect-all webview`：收集 pywebview 的运行时资源（含 WebView2Loader）
- `--add-data "webgui;webgui"`：把界面（HTML/CSS/JS）打进 exe
- 运行机器需有 **WebView2 Runtime**（Win10/11 系统自带，一般无需安装）

生成的 `dist/ExcelSplitter.exe` 即为可分发单文件。

> 说明：本仓库**不收录** `dist/` 下的 exe 二进制（避免长期占用仓库体积）。按上面命令即可在任何机器重新生成，结果一致。

## 发布新版本（GitHub Actions 自动打包）

仓库已内置 `.github/workflows/build-release.yml`：推送一个 `v*` 格式的 tag 时，GitHub 会在云端 Windows 运行器上自动用 PyInstaller 打包 exe，并创建 Release、把**带版本号的 exe**（如 `ExcelSplitter-v2.0.0.exe`）作为下载附件。

```bash
# 1. 确保所有改动已提交并推送到 GitHub
git add -A && git commit -m "release: v2.1.0" && git push origin main

# 2. 打 tag 并推送（push tag 即触发 Actions 自动构建 + 发版）
git tag v2.1.0
git push origin v2.1.0
```

推送 tag 后，到仓库的 **Actions** 页看构建进度；成功后 **Releases** 页会出现对应版本，里面带可直接下载的 `ExcelSplitter-<版本号>.exe`。

> 改代码后发新版：本地重新 `git tag v2.1.0`（递增版本号）→ `git push origin v2.1.0` 即可，无需手动上传文件。

## Gitee 镜像（源码自动同步）

GitHub 为主仓库，代码会**自动同步**到 Gitee 镜像（[`gitee.com/wiggins-kong/ExcelSplitter`](https://gitee.com/wiggins-kong/ExcelSplitter)）：你 push 到 GitHub 后，`.github/workflows/sync-to-gitee.yml` 会自动把全部分支和 tags 推送到 Gitee，全程云端、无需任何本地操作。

**Release 附件（exe）不会云端自动同步**——云端构建机跨境上传大文件到 Gitee 稳定失败，因此用**本地脚本镜像**：GitHub 发版后在本机跑一条命令（token 放环境变量 `GITEE_TOKEN`，见 `DEVELOPMENT.md` 坑 11）：

```bash
python scripts/publish_gitee_release.py v2.1.0
```

脚本会自动复制 GitHub Release 的标题和说明、在 Gitee 建同名 Release 并直连上传 exe（省略 exe 参数时优先取 `dist/` 下现成文件，否则从 GitHub Release 下载）。

## 目录结构

```
excel_splitter.py            # 核心逻辑 + CLI 入口（读取层 / 拆分逻辑，无 GUI 依赖）
excel_splitter_gui_web.py    # GUI（pywebview + WebView2 + 真 Mica + 深浅主题，被 excel_splitter.py 无参数调用）
webgui/index.html            # GUI 界面（HTML/CSS/JS，Win11 Fluent 设计，双主题 Tokens）
dpi_manifest.xml             # 高 DPI 感知清单（PerMonitorV2），打包时嵌入 exe
assets/icons/                # 项目 Logo 与多尺寸 .ico（图标源文件 + make_ico.py 生成脚本）
.github/workflows/           # GitHub Actions：build-release（发版打包）+ sync-to-gitee（源码同步 Gitee）
CHANGELOG.md                 # 版本变更记录
DEVELOPMENT.md               # 开发进度与踩坑记录（换机/新会话接续用）
design/                      # UI 设计稿（HTML demo，含深色主题初稿）
.gitignore
README.md
```

## 图标资源

- `assets/icons/ExcelSplitter.ico`：打包用的多尺寸图标（16/24/32/48/64/128/256，含 PNG 压缩帧，圆角外透明）。
- `assets/icons/make_ico.py`：从设计源图重新生成 `.ico` 的脚本（需要 Pillow）：
  `python assets/icons/make_ico.py`
- 其余 `icon-*.jpg/png` 为设计源图（白底 / 深底 / 纯图形 / 透明底，2048×2048），`logo-candidate-*.jpg` 为首轮候选稿。

## 与原 WPS 宏的关系

本工具还原并改进了 WPS 表格的 JS 宏 `SplitByNonAdjacentSelection`：

- 不再依赖 WPS / Excel 正在运行，纯读文件拆分
- 修正了原宏的一个文件名 bug：原宏用 `_||_` 作分隔符，但在文件名替换时正则写错（`/_||_/g` 实际被解析成「_ 或 | 或 _」），会导致文件名冒出一堆空格。本工具改用**元组作分组键**、文件名用 ` - ` 连接，干净且无歧义
- 数据清洗逻辑（空值 / 空格 / 换行处理）原样保留

## License

MIT
