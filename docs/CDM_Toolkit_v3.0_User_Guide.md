# CDM Toolkit v3.0 使用说明文档

**版本**: 3.0.0  
**日期**: 2026-05-30  
**适用对象**: DM（数据管理）数据管理员

---

## 目录

1. [概述](#1-概述)
2. [环境准备](#2-环境准备)
3. [Web 工具使用指南](#3-web-工具使用指南)
4. [VBA 宏使用指南](#4-vba-宏使用指南)
5. [Python 引擎使用指南](#5-python-引擎使用指南)
6. [VBA ↔ Web 互操作](#6-vba--web-互操作)
7. [典型工作流示例](#7-典型工作流示例)
8. [常见问题](#8-常见问题)

---

## 1. 概述

CDM Toolkit v3.0 是一个面向临床数据管理（CDM）的三层架构工具包：

```
┌─────────────────────────────────────────┐
│  VBA 宏 (Excel内)  │  Web UI (浏览器)    │  ← 表示层
├─────────────────────────────────────────┤
│          cdm_engine (Python 包)          │  ← 逻辑层
├─────────────────────────────────────────┤
│   Excel (.xlsx)  │  SAS (.sas7bdat)  │  CSV  │  ← 数据层
└─────────────────────────────────────────┘
```

### 什么时候用什么？

| 场景 | 用什么 | 怎么用 |
|------|--------|--------|
| 在 Excel 里审数据，想快速替换术语 | **VBA** | `Alt+F8` → `DM_TextStandardise` |
| 要合并 20 个 Site 发来的 Excel | **Web** | 浏览器拖拽上传 → 一键合并 |
| 想批量处理几百个文件的日期列 | **Web** | 上传 → 配置参数 → 下载结果 |
| 需要比较两个版本的数据差异 | **VBA** | `DM_TrackChanges` 在原 sheet 上标注 |
| 扫描特殊字符准备 EDC Query | **Web** | 上传 → 扫描 → 导出 Rave 报告 |
| 写脚本自动化批量处理 | **Python** | `from cdm_engine import ...` |

---

## 2. 环境准备

### 2.1 安装 Python 依赖

```bash
# 安装 CDM Toolkit（开发模式）
cd e:\Project\Tools\Tools_DM
pip install -e .

# 如果需要 SAS 数据集支持
pip install pyreadstat
```

### 2.2 安装 VBA 宏到 Excel

**方法一：手动导入（推荐）**

1. 打开 Excel
2. 按 `Alt+F11` 进入 VBA 编辑器
3. 菜单：`File` → `Import File`（或 `文件` → `导入文件`）
4. 进入 `vba/modules/` 目录，选中所有 `.bas` 文件（按住 Ctrl 多选），点"打开"
5. 再导入 `vba/CDM_Toolkit.bas`
6. 再导入 `vba/CDM_Install.bas`
7. 关闭 VBA 编辑器，回到 Excel

**方法二：安装到 Personal.xlsb（所有 Excel 文件都能用）**

1. 先按方法一手动导入所有 `.bas` 到当前工作簿
2. 按 `Alt+F8`，找到 `DM_InstallAll`，点击"执行"
3. 所有宏被复制到 `PERSONAL.XLSB`，之后任意 Excel 文件按 `Alt+F8` 都能看到

> **注意**: 使用 `DM_InstallAll` 需要在 Excel 中启用"信任对 VBA 工程对象模型的访问"：
> `文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选"信任对 VBA 工程对象模型的访问"`

---

## 3. Web 工具使用指南

### 3.1 启动 Web 工具

**方式一：命令行启动**

```bash
cd e:\Project\Tools\Tools_DM
python server.py
```

浏览器自动打开 `http://localhost:8520`

**方式二：从 Excel 启动**

在 Excel 中按 `Alt+F8` → 选择 `DM_LaunchWeb` → 执行

**自定义端口：**

```bash
python server.py --port 9999        # 使用 9999 端口
python server.py --no-browser       # 不自动打开浏览器
```

### 3.2 仪表盘

启动后看到的主页就是仪表盘，展示所有功能入口：

- **上半部分**（蓝色 Web 标签）：6 个 Web 端功能，点击卡片进入
- **下半部分**（绿色 VBA 标签）：6 个 VBA 端功能，提示在 Excel 中使用
- **底部状态栏**：显示引擎版本和 VBA 安装状态

### 3.3 工作簿合并

**适用场景**：多个 Site 发来的 Excel 需要合并成一个工作簿

**操作步骤**：

1. 点击首页"工作簿合并"卡片，或顶部导航栏"工作簿合并"
2. 将 `.xlsx` 或 `.xls` 文件拖拽到上传区域（或点击选择）
3. 已选文件会显示在列表中，可点击 ✕ 移除
4. 勾选选项：
   - `自动生成 TOC 目录`（推荐勾选）
   - `文件名作为 sheet 前缀`（避免 sheet 重名）
5. 点击"开始合并"
6. 看到进度条，完成后自动下载 `merged.xlsx`

**注意事项**：
- 每个源文件的 sheet 名会加上文件名前缀（如 `SITE01_AE`）
- Excel 的 31 字符 sheet 名限制会自动处理
- 支持混合 `.xlsx` 和 `.xls` 格式

### 3.4 工作表拆分

**适用场景**：一个大的 DM 数据集需要按 Site/项目拆分到不同 sheet

**操作步骤**：

1. 点击"工作表拆分"卡片
2. 上传一个 `.xlsx` 文件
3. 填写参数：
   - **拆分列**：按哪一列的值拆（如 `A`、`B`、`C`）
   - **表头行数**：保留几行表头复制到每个子 sheet（默认 `1`）
4. 点击"开始拆分"
5. 自动下载 `split.xlsx`

**示例**：DM 数据 A 列是 `SITEID`，拆分列填 `A`，表头行填 `1` → 每个 Site 一个 sheet

### 3.5 特殊字符扫描

**适用场景**：提交 EDC 前检查数据中的非标准字符（如 °、↑、ï 等）

**操作步骤**：

1. 点击"特殊字符扫描"卡片
2. 上传 Excel 文件
3. 选择 EDC 系统：
   - `Rave` — 生成 Rave 格式报告
   - `Clinflash` — 生成 Clinflash 格式报告
4. 点击"开始扫描"
5. 查看结果表格（Sheet 名、单元格地址、内容）
6. 点击"导出 JSON 报告"保存结果

**默认允许的字符**：英文字母、数字、空格、`-` `_` `.` `,` `:` `;` `(` `)` `/`

### 3.6 文件目录浏览

**适用场景**：需要快速浏览项目文件夹内容并导出清单

**操作步骤**：

1. 点击"文件目录浏览"卡片
2. 输入文件夹路径（如 `C:\Data\CDM\Study001`）
3. 选择文件类型过滤（所有文件 / Excel / SAS / CSV / PDF）
4. 勾选"包含子文件夹"（默认勾选）
5. 点击"浏览"
6. 树形列表展示所有文件
7. 点击"导出 JSON"保存文件清单

### 3.7 日期计算

**适用场景**：批量处理日期列（如 AE 开始日期 +30 天）

**操作步骤**：

1. 点击"日期计算"卡片
2. 上传 Excel 文件
3. 填写参数：
   - **目标列**：要处理的日期列字母（如 `F`）
   - **加减天数**：正数加、负数减（如 `30` 或 `-7`）
   - **日期格式**：结果格式（默认 `YYYY-MM-DD`）
4. 点击"执行"
5. 自动下载 `dates_result.xlsx`

**注意**：非日期格式的单元格会被跳过，不会报错。

### 3.8 单位转换

**适用场景**：将实验室检验结果统一转换为标准单位

**操作步骤**：

1. 点击"单位转换"卡片
2. 上传 Excel 文件（含检验数据）
3. 填写参数：
   - **检验项目列**：如 `D`（LBTEST）
   - **结果值列**：如 `E`（LBORRES）
   - **单位列**：如 `F`（LBORRESU）
   - **目标标准单位**：如 `mg/dL`
   - **检验项目筛选**：留空=全部转换，或填入 `Glucose` 只转换该检验
4. 点击"执行转换"
5. 自动下载 `units_result.xlsx`

**内置转换因子**（24组）：

| 类别 | 示例转换 |
|------|---------|
| 质量 | g↔mg (×1000), mg↔ug (×1000), kg↔g (×1000) |
| 体积 | L↔mL (×1000), dL↔L (×0.1) |
| 长度 | cm↔mm (×10), m↔cm (×100) |
| 浓度 | mg/dL↔g/L (×0.01), umol/L↔mg/dL (×0.0113) |
| 时间 | h↔min (×60), d↔h (×24) |

---

## 4. VBA 宏使用指南

### 4.1 模块总览

VBA 工具箱共 **8 个核心模块**，按功能分类：

| 分类 | 模块 | 宏名 | 功能 |
|------|------|------|------|
| 数据清理 | modTextStandardise | `DM_TextStandardise` | CDM 术语标准化 |
| | modHighlightDup | `DM_HighlightDuplicates` | 连续重复值高亮 |
| | modBlankRows | `DM_InsertBlankEveryN` | 每N行插入空白行 |
| | | `DM_InsertBlankBetweenGroups` | 组间插入空白行 |
| | | `DM_DeleteBlankRows` | 删除完全空白行 |
| 文档管理 | modTOC | `DM_GenerateTOC` | 生成带超链接的目录 |
| | | `DM_GenerateDeriveTOC` | SAS Derive 风格目录 |
| 工作表 | modSheetManager | `DM_ListSheetNames` | 列出所有 sheet 名称 |
| | | `DM_BatchRenameSheets` | 批量重命名 sheet |
| | | `DM_DeleteHiddenSheets` | 删除隐藏 sheet |
| 审核追溯 | modTrackChanges | `DM_TrackChanges` | 与原版比较并标记修改 |
| | | `DM_ClearTrackChanges` | 清除修改标记 |
| | | `DM_ChangesReport` | 生成结构化修改报告 |
| 计算 | modDateCalc | `DM_AddDaysToColumn` | 日期列加减天数 |
| | | `DM_CalcDateDiff` | 计算两列日期差 |
| | | `DM_DateAddSelection` | 选中日期加天数 |
| | modUnitConvert | `DM_ConvertLabUnits` | 实验室单位转换 |
| 工具 | CDM_Install | `DM_LaunchWeb` | 一键启动 Web 工具 |

### 4.2 数据清理

#### DM_TextStandardise — 文本术语标准化

**功能**：批量替换非标准 CDM 术语为标准术语

**内置替换规则**：
| 旧术语 | 新术语 |
|--------|--------|
| Recovered | Recovered / Resolved |
| Ongoing | Not Recovered / Not Resolved |
| Reasonable Possib | Reasonable Possibility |
| Not Recovered/Not Resolved | Not Recovered / Not Resolved |
| Recovered/Resolved | Recovered / Resolved |

**使用**：
```vb
' 使用内置映射表
DM_TextStandardise ActiveSheet

' 使用自定义映射表
Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
map.Add "AE", "Adverse Event"
DM_TextStandardise ActiveSheet, map
```

#### DM_HighlightDuplicates — 重复值高亮

**功能**：高亮相邻行中相同列值连续重复的单元格（黄色）

**使用**：
```vb
' 默认 G-K 列（7-11），从第2行开始
DM_HighlightDuplicates ActiveSheet

' 只检查 E 列（第5列），红色标记
DM_HighlightDuplicates ActiveSheet, 5, 5, 2, vbRed

' 单列便捷方法
DM_HighlightDupSingleCol ActiveSheet, 5
```

#### DM_InsertBlankEveryN — 每N行插入空白行

**功能**：每隔 N 行插入一行空白（自底向上）

**使用**：
```vb
' 每5行插入一行
DM_InsertBlankEveryN ActiveSheet, 5, "A"

' 每10行插入一行
DM_InsertBlankEveryN ActiveSheet, 10
```

#### DM_InsertBlankBetweenGroups — 组间插入空白行

**功能**：当指定列的值变化时在组间插入空白行

**使用**：
```vb
' 当A列的值变化时插入空白行
DM_InsertBlankBetweenGroups ActiveSheet, "A"
```

#### DM_DeleteBlankRows — 删除空白行

**功能**：删除工作表中所有完全空白的行

**使用**：
```vb
DM_DeleteBlankRows ActiveSheet
```

### 4.3 文档管理

#### DM_GenerateTOC — 生成目录

**功能**：创建/重建目录 sheet，包含指向所有可见 sheet 的超链接，并在每个被链接的 sheet 上添加返回箭头

**使用**：
```vb
' 默认使用 "TOC" 作为目录名
DM_GenerateTOC ActiveWorkbook

' 自定义目录名 + 排除特定 sheet
DM_GenerateTOC ActiveWorkbook, "目录", "Settings,Temp"
```

#### DM_GenerateDeriveTOC — SAS Derive 风格目录

**功能**：A 列＝超链接 sheet 名称，B 列＝从指定单元格提取的描述

**使用**：
```vb
' 从每个 sheet 的 V2 单元格提取描述
DM_GenerateDeriveTOC ActiveWorkbook, "V2", "TOC"
```

### 4.4 工作表管理

#### DM_ListSheetNames — 列出 sheet 名称

```vb
Dim names() As String
names = DM_ListSheetNames(ActiveWorkbook)        ' 仅可见
names = DM_ListSheetNames(ActiveWorkbook, True)  ' 包含隐藏
```

#### DM_BatchRenameSheets — 批量重命名

```vb
Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
map.Add "Sheet1", "AE_Data"
map.Add "Sheet2", "LB_Data"
DM_BatchRenameSheets ActiveWorkbook, map
```

#### DM_DeleteHiddenSheets — 删除所有隐藏 sheet

```vb
Dim deleted() As String
deleted = DM_DeleteHiddenSheets(ActiveWorkbook)
```

### 4.5 审核追溯

#### DM_TrackChanges — 修改痕迹跟踪

**功能**：比较当前工作表与原始备份副本：
- 修改过的单元格 → 绿色填充 + 红色边框
- 添加批注显示 `旧值: X → 新值: Y`

**使用**：
```vb
' 完整标记（批注 + 颜色高亮）
DM_TrackChanges ActiveSheet, "C:\backup\original.xlsx"

' 仅批注，不改变颜色
DM_TrackChanges ActiveSheet, "C:\backup\original.xlsx", True, False
```

#### DM_ClearTrackChanges — 清除标记

```vb
DM_ClearTrackChanges ActiveSheet
```

#### DM_ChangesReport — 生成修改报告

**功能**：创建新的"修改报告"sheet，结构化列出所有变更：
- A 列：单元格地址
- B 列：变更类型（新增/删除/修改）
- C 列：旧值
- D 列：新值

```vb
DM_ChangesReport ActiveSheet, "C:\backup\original.xlsx"
```

### 4.6 计算工具

#### DM_AddDaysToColumn — 日期列加减天数

```vb
' F列所有日期 +30天
DM_AddDaysToColumn ActiveSheet, "F", 30

' G列所有日期 -7天，格式 DD/MM/YYYY
DM_AddDaysToColumn ActiveSheet, "G", -7, 2, "DD/MM/YYYY"
```

#### DM_CalcDateDiff — 计算日期差

```vb
' (G列 - F列) 的天数差 写入 M列
DM_CalcDateDiff ActiveSheet, "F", "G", "M"
```

#### DM_DateAddSelection — 选中日期加减

```vb
' 选中一批日期单元格，然后运行
DM_DateAddSelection 30     ' 全部 +30天
```

#### DM_ConvertLabUnits — 实验室单位转换

```vb
' 将 Glucose 检验结果统一转换为 mg/dL
' D列=检验项目, E列=结果值, F列=单位
DM_ConvertLabUnits ActiveSheet, "D", "E", "F", "mg/dL", "Glucose"

' 全部检验项目转换
DM_ConvertLabUnits ActiveSheet, "D", "E", "F", "mg/dL"
```

### 4.7 工具

#### DM_LaunchWeb — 启动 Web 工具

**功能**：一键启动本地 CDM Web 服务器并打开浏览器

```vb
DM_LaunchWeb
```

前提条件：
- Python 3.10+ 已安装
- `server.py` 和 `cdm_engine/` 在同一目录下
- 已安装依赖：`pip install fastapi uvicorn[standard] python-multipart`

---

## 5. Python 引擎使用指南

### 5.1 基础导入

```python
from openpyxl import load_workbook
from cdm_engine import *

# 查看版本
from cdm_engine import __version__
print(__version__)  # 3.0.0
```

### 5.2 完整可用函数列表

| 函数 | 功能 | 返回 |
|------|------|------|
| `standardise_terms(ws, term_map=None)` | 术语标准化 | `int` (替换次数) |
| `highlight_duplicates(ws, col_start, col_end, row_start, fill)` | 重复高亮 | `None` |
| `insert_blank_every_n(ws, n, col_letter)` | 每N行插空行 | `int` |
| `insert_blank_between_groups(ws, col_letter)` | 组间插空行 | `int` |
| `delete_blank_rows(ws)` | 删除空白行 | `int` |
| `generate_toc(wb, toc_name, exclude_sheets)` | 生成目录 | `Worksheet` |
| `generate_sas_derive_toc(wb, desc_col, toc_name)` | SAS风格目录 | `Worksheet` |
| `split_sheet_by_column(wb, sheet, col, header_rows)` | 按列拆分 | `list[str]` |
| `merge_workbooks(paths, output, create_toc)` | 合并工作簿 | `Workbook` |
| `track_changes(ws, original_path)` | 标记修改 | `int` |
| `changes_report(ws, original_path)` | 修改报告 | `list[dict]` |
| `add_days_to_column(ws, col, days, row_start, fmt)` | 日期加天数 | `int` |
| `calc_date_diff(ws, c1, c2, result, row_start)` | 日期差计算 | `int` |
| `convert_lab_units(ws, test, result, unit, std, filter)` | 单位转换 | `int` |
| `batch_rename_sheets(wb, name_map)` | 批量重命名 | `None` |
| `list_sheet_names(wb)` | 列出sheet名 | `list[str]` |
| `delete_hidden_sheets(wb)` | 删除隐藏sheet | `list[str]` |
| `scan_special_chars(ws, pattern)` | 扫描特殊字符 | `list[tuple]` |
| `generate_file_listing(folder, output, ...)` | 文件目录Excel | `Workbook` |
| `generate_folder_tree(folder, output)` | 文件夹树Excel | `Workbook` |
| `apply_header_style(ws, row)` | 应用表头样式 | `None` |
| `auto_width(ws, max_w, min_w)` | 自动列宽 | `None` |

### 5.3 常用脚本示例

#### 示例1：AE 数据清理流水线

```python
from openpyxl import load_workbook
from cdm_engine import *

wb = load_workbook("AE_Raw.xlsx")
ws = wb["AE"]

# 1. 标准化术语
n = standardise_terms(ws)
print(f"术语替换: {n} 处")

# 2. 高亮 AETERM 列重复值
highlight_duplicates(ws, col_start=5, col_end=5, row_start=2)

# 3. 删除空白行
d = delete_blank_rows(ws)
print(f"删除空白行: {d} 行")

# 4. 按 SUBJID 分组插空行
insert_blank_between_groups(ws, "C")

# 5. 计算 AE 持续天数 (G-F)
ws.cell(row=1, column=13, value="AEDUR")
calc_date_diff(ws, "F", "G", "M")

# 6. 生成目录
generate_toc(wb)

# 7. 自动列宽
for s in wb.worksheets:
    auto_width(s)

wb.save("AE_Cleaned.xlsx")
print("完成！")
```

#### 示例2：多 Site 数据合并

```python
from cdm_engine import merge_workbooks
import glob

# 收集所有 Site 文件
files = glob.glob("C:/Data/Study001/SITE*.xlsx")

# 合并为一个工作簿
merged = merge_workbooks(files, "Study001_All.xlsx", create_toc=True)
print(f"合并完成: {len(merged.sheetnames)} 个 sheet")
```

#### 示例3：实验室数据标准化

```python
from openpyxl import load_workbook
from cdm_engine import convert_lab_units, auto_width

wb = load_workbook("Lab_Raw.xlsx")
ws = wb["LB"]

# 统一 Glucose 为 mg/dL
n1 = convert_lab_units(ws, "D", "E", "F", "mg/dL", test_filter="Glucose")
print(f"Glucose 转换: {n1} 行")

# 统一 Hemoglobin 为 g/dL
n2 = convert_lab_units(ws, "D", "E", "F", "g/dL", test_filter="Hemoglobin")
print(f"Hemoglobin 转换: {n2} 行")

# 统一 Creatinine 为 mg/dL
n3 = convert_lab_units(ws, "D", "E", "F", "mg/dL", test_filter="Creatinine")
print(f"Creatinine 转换: {n3} 行")

auto_width(ws)
wb.save("Lab_Standardized.xlsx")
```

#### 示例4：SAS 数据集转 Excel

```python
from cdm_engine.sas_utils import sas_to_excel, read_sas

# 方式1：直接转换
sas_to_excel("ae.sas7bdat", "AE_From_SAS.xlsx")

# 方式2：先读取为 DataFrame 再处理
df = read_sas("lb.sas7bdat")
print(f"LB 数据集: {len(df)} 行, {len(df.columns)} 列")
print(df.head())
```

#### 示例5：批量文件目录生成

```python
from cdm_engine import generate_file_listing, generate_folder_tree

# 生成带超链接的文件清单
generate_file_listing(
    folder_path="C:/Data/Study001",
    output_path="File_Listing.xlsx",
    include_subfolders=True,
    file_pattern="*.xlsx"
)

# 生成树形文件夹结构
generate_folder_tree(
    folder_path="C:/Data/Study001",
    output_path="Folder_Tree.xlsx"
)
```

---

## 6. VBA ↔ Web 互操作

### 三种启动方式

| 从 | 操作 | 结果 |
|----|------|------|
| **命令行** | `cdm-tools serve` 或 `python server.py` | 浏览器打开 `http://localhost:8520` |
| **Excel VBA** | `Alt+F8` → `DM_LaunchWeb` | 同上 |
| **Python** | `server.start_server(port=8520)` | 同上 |

### 典型互操作流程

```
Excel (VBA)                     Web (浏览器)                     Excel (VBA)
   │                                │                                │
   ├─ Alt+F8 → DM_TextStandardise  │                                │
   ├─ 术语标准化完成                 │                                │
   ├─ DM_HighlightDuplicates       │                                │
   ├─ 重复值高亮完成                 │                                │
   │                                │                                │
   ├─ DM_LaunchWeb ──────────────→ 浏览器打开                        │
   │                                ├─ 上传多个Site文件               │
   │                                ├─ 合并 + 生成TOC                 │
   │                                ├─ 下载 merged.xlsx              │
   │                                │                                │
   │  ← 在Excel中打开 merged.xlsx ──┘                                │
   ├─ DM_TrackChanges 比较版本      │                                │
   ├─ 审核完成                      │                                │
```

---

## 7. 典型工作流示例

### 工作流 A：收到 CRO 数据后的日常清理

**场景**：CRO 发来 AE 和 LB 的 Excel，需要标准化术语、检查重复、清理空白行

```
1. 打开 AE.xlsx
2. Alt+F8 → DM_TextStandardise       → 标准化术语
3. Alt+F8 → DM_HighlightDuplicates   → 高亮重复
4. Alt+F8 → DM_DeleteBlankRows       → 删除空行
5. 同上处理 LB.xlsx
6. Alt+F8 → DM_GenerateTOC           → 生成目录
7. 保存
```

### 工作流 B：多 Site 数据汇聚

**场景**：20 个 Site 各自提交了 Excel，需要合并成一份

```
1. Alt+F8 → DM_LaunchWeb            → 打开 Web 工具
2. 点击"工作簿合并"
3. 拖入 20 个 Excel 文件
4. 勾选"自动生成 TOC 目录" + "文件名作为 sheet 前缀"
5. 点击"开始合并"
6. 下载 merged.xlsx
7. 在 Excel 中打开 → Alt+F8 → DM_GenerateTOC → 审核
```

### 工作流 C：EDC 提交前特殊字符检查

**场景**：数据要导入 Rave，需要检查是否有特殊字符会导致导入失败

```
1. Alt+F8 → DM_LaunchWeb
2. 点击"特殊字符扫描"
3. 上传 Excel 文件
4. 选择 EDC 系统 "Rave"
5. 点击"开始扫描"
6. 查看结果表格 → 定位有问题的单元格
7. 回到 Excel → 修复特殊字符
8. 重新扫描确认无问题
9. 导出 Rave 报告，提交 Query
```

### 工作流 D：Python 脚本自动化

**场景**：每周需要做同样的数据处理，写脚本自动执行

```python
# weekly_process.py — 每周数据处理脚本
from openpyxl import load_workbook
from cdm_engine import *
import glob

DATA_DIR = r"C:\Data\Study001\Weekly"
OUT_DIR  = r"C:\Data\Study001\Processed"

# 1. 合并本周所有文件
files = glob.glob(f"{DATA_DIR}/Week*.xlsx")
merge_workbooks(files, f"{OUT_DIR}/Weekly_Merged.xlsx")

# 2. 清理
wb = load_workbook(f"{OUT_DIR}/Weekly_Merged.xlsx")
for ws in wb.worksheets:
    standardise_terms(ws)
    delete_blank_rows(ws)
    apply_header_style(ws)
    auto_width(ws)

# 3. 统一单位
convert_lab_units(wb["LB"], "D", "E", "F", "mg/dL")

# 4. 输出
generate_toc(wb)
wb.save(f"{OUT_DIR}/Weekly_Cleaned.xlsx")
print("本周数据处理完成！")
```

---

## 8. 常见问题

### Q: Web 工具启动后浏览器没有自动打开？

手动打开浏览器，访问 `http://localhost:8520`

### Q: 端口 8520 被占用？

使用自定义端口：
```bash
python server.py --port 9999
```

然后访问 `http://localhost:9999`

### Q: VBA 宏运行时提示"找不到工程或库"？

在 VBA 编辑器中：`Tools` → `References` → 查找标有"MISSING"的引用 → 取消勾选 → 确定

### Q: DM_LaunchWeb 提示找不到 server.py？

确保 Excel 工作簿和 `server.py` 在同一个目录下。或手动在命令行启动：
```bash
cd e:\Project\Tools\Tools_DM
python server.py
```

### Q: 特殊字符扫描结果太多，怎么筛选？

结果以 JSON 格式展示，点击"导出 JSON 报告"后在 Excel 或文本编辑器中打开，按 Sheet 筛选。

### Q: 可以批量处理多个文件的特殊字符扫描吗？

目前的 Web 界面每次处理一个文件。如需批量处理，使用 Python 脚本：
```python
import glob
from openpyxl import load_workbook
from cdm_engine import scan_special_chars

for fpath in glob.glob("*.xlsx"):
    wb = load_workbook(fpath)
    for ws in wb.worksheets:
        findings = scan_special_chars(ws)
        if findings:
            print(f"{fpath} / {ws.title}: {len(findings)} 处特殊字符")
```

### Q: 如何添加自定义单位转换因子？

在 VBA 中：
```vb
DM_AddUnitConversion "ng", "mg", 0.000001
```

在 Python 中：
```python
from cdm_engine.conversions import UNIT_CONVERSIONS
UNIT_CONVERSIONS[("ng", "mg")] = 0.000001
```

### Q: Web 工具能在无网络环境下使用吗？

可以。Web 工具运行在 `localhost`，完全离线，不需要互联网连接。

### Q: 数据会传到外部服务器吗？

不会。所有处理都在本地完成：
- Web 工具：`localhost` 本地服务，文件不上传外部
- VBA 宏：Excel 内运行
- Python 引擎：本地文件系统读写

### Q: 企业 IT 策略阻止了 localhost 端口怎么办？

1. 尝试更换端口：`python server.py --port 8080`
2. 只在 Excel 中用 VBA 宏（不需要网络端口）
3. 直接在 Python 脚本中调用 `cdm_engine`（也不需要端口）

---

## 版本历史

| 版本 | 日期 | 变更 |
|------|------|------|
| 3.0.0 | 2026-05-30 | 三层架构重新设计：VBA精简+Web前端+Python统一引擎 |
| 2.x | 2026-05-29 | VBA全量实现（12模块）+ Python双文件 |
| 1.x | 2026-05-28 | 初始Python实现（cdm_toolkit.py 14函数） |
