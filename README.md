# Excel 按列拆分工具（ExcelSplitter）

把一张 Excel 总表，按你**指定的若干列**自动拆分成多个独立的 `.xlsx` 文件。

这是 WPS 表格 JS 宏 `SplitByNonAdjacentSelection` 的**独立桌面版实现**——不依赖 WPS，双击即用，跨平台（Windows / 打包后单文件分发）。

## 功能特性

- **按多列组合分组**：选中「地区」「产品」两列 → 自动拆成 `华东 - 苹果.xlsx`、`华北 - 香蕉.xlsx`……每个唯一组合生成一个文件
- **强力数据清洗**：空值 → `空白`、去首尾空格、去单元格内换行 / 制表符（专门解决「看着一样却没分到一组」的坑）
- **非破坏性**：原表只读不改，仅新建 / 保存拆分结果
- **完整保留**：每个拆分文件含完整表头与所有列
- **非法文件名处理**：`\ / * ? " < > |` → `-`、冒号替换、长度截断到 150
- **高分屏适配**：窗口声明 DPI 感知，高缩放下界面清晰不模糊

## 使用方法

### 图形界面（推荐）

1. 双击 `dist/ExcelSplitter.exe`（需自行构建，见下方）
2. 「浏览」选择 Excel 文件（`.xlsx` / `.xlsm`）
3. 选择工作表、确认表头行（默认第 1 行）
4. 在「拆分依据列」里用 **Ctrl / Shift** 多选要按哪些列拆
5. 选择输出文件夹（默认在原文件旁建一个同名子目录）
6. 点「开始拆分」→ 带进度条与运行日志

### 命令行

```bash
ExcelSplitter.exe 文件.xlsx --cols 1,3 --out 输出目录 --header 1
```

| 参数 | 说明 |
| --- | --- |
| `input` | 输入 xlsx 路径（位置参数） |
| `--cols` | 拆分列（1-based，逗号分隔，如 `1,3`） |
| `--sheet` | 指定工作表名 |
| `--header` | 表头行号，默认 `1` |
| `--out` | 输出目录 |

## 从源码构建

需要 **Windows + Python 3.10+**（建议系统自带 Python，需含 `tkinter`）。

```bash
pip install openpyxl pyinstaller
pyinstaller --onefile --windowed --name ExcelSplitter --noupx --manifest dpi_manifest.xml excel_splitter.py
```

生成的 `dist/ExcelSplitter.exe` 即为可分发单文件。

> 说明：本仓库**不收录** `dist/` 下的 exe 二进制（避免长期占用仓库体积）。按上面命令即可在任何机器重新生成，结果一致。

## 发布新版本（GitHub Actions 自动打包）

仓库已内置 `.github/workflows/build-release.yml`：推送一个 `v*` 格式的 tag 时，GitHub 会在云端 Windows 运行器上自动用 PyInstaller 打包 exe，并创建 Release、把 `ExcelSplitter.exe` 作为下载附件。

```bash
# 1. 确保所有改动已提交并推送到 GitHub
git add -A && git commit -m "release: v1.0.0" && git push origin main

# 2. 打 tag 并推送（push tag 即触发 Actions 自动构建 + 发版）
git tag v1.0.0
git push origin v1.0.0
```

推送 tag 后，到仓库的 **Actions** 页看构建进度；成功后 **Releases** 页会出现 `v1.0.0`，里面带可直接下载的 `ExcelSplitter.exe`。

> 改代码后发新版：本地重新 `git tag v1.0.1`（递增版本号）→ `git push origin v1.0.1` 即可，无需手动上传文件。

## 目录结构

```
excel_splitter.py   # 主程序源码（tkinter GUI + openpyxl 拆分逻辑）
dpi_manifest.xml    # 高 DPI 感知清单（PerMonitorV2），打包时嵌入 exe
.gitignore
README.md
```

## 与原 WPS 宏的关系

本工具还原并改进了 WPS 表格的 JS 宏 `SplitByNonAdjacentSelection`：

- 不再依赖 WPS / Excel 正在运行，纯读文件拆分
- 修正了原宏的一个文件名 bug：原宏用 `_||_` 作分隔符，但在文件名替换时正则写错（`/_||_/g` 实际被解析成「_ 或 | 或 _」），会导致文件名冒出一堆空格。本工具改用**元组作分组键**、文件名用 ` - ` 连接，干净且无歧义
- 数据清洗逻辑（空值 / 空格 / 换行处理）原样保留

## License

MIT
