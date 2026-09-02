# Changelog

本项目的所有重要变更都记录在此文件中。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

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