# CDM VBA & Python 工具包 — 设计规格书

**日期**: 2026-05-30  
**状态**: 已批准  
**需求来源**: reference/宏目录-DM1.xlsx、reference/DMⅡ常用宏.xlsx、reference/在excel中使用VB宏-DMI.docx

---

## 1. 概述

基于 reference 文件夹中 Excel 需求文档，开发一套完整的、生产级的 VBA 宏工具包，用于临床数据管理（CDM）Excel 工作流。VBA 工具包与已有的 `cdm_toolkit.py` Python 实现完全对齐。同时，根据 VBA 分析中发现的缺口，补充 Python 工具包缺失的功能。

### 目标

- **12 个 VBA 模块**（.bas 文件），覆盖 DMⅡ常用宏.xlsx 中记载的所有 VBA Excel 宏需求
- **1 个 Python 补充文件**（`cdm_toolkit_plus.py`），填补已有 Python 工具包的功能缺口
- **双向完全对齐**：每个 VBA 函数都有对应的 Python 实现，反之亦然
- **生产级质量**：统一命名规范、完整参数化、错误处理、中英双语注释、使用示例

### 不在范围内

- SAS 宏（已在 `DM_Toolbox.sas` 和 `sas_macros/` 中实现）
- VBA 用户窗体 / GUI 对话框（纯宏模块）
- Excel 加载项打包（`.xlam`）— 交付物为 .bas 模块文件

---

## 2. 目录架构

```
Tools_DM/
├── cdm_toolkit.py              # [已有] Python CDM工具包 — 14个函数
├── cdm_toolkit_plus.py         # [新建] Python补充 — 文件目录生成、增强修改痕迹
├── DM_Toolbox.sas              # [已有] SAS宏工具包主文件
├── vba/                         # [新建] VBA模块目录
│   ├── CDM_Toolkit.bas          #   主模块 — 汇总所有子模块（用于批量导入）
│   ├── CDM_Install.bas          #   一键安装：将所有宏注册到 Personal.xlsb
│   ├── README.md                #   中英双语VBA使用说明书
│   └── modules/                 #   独立 .bas 文件（可单独导入）
│       ├── modTextStandardise.bas   ' 文本术语标准化
│       ├── modHighlightDup.bas      ' 重复值高亮
│       ├── modBlankRows.bas         ' 空白行管理
│       ├── modTOC.bas               ' 目录生成
│       ├── modSheetSplit.bas        ' 工作表拆分
│       ├── modWorkbookMerge.bas     ' 工作簿合并
│       ├── modSheetManager.bas      ' 工作表管理
│       ├── modTrackChanges.bas      ' 修改痕迹跟踪
│       ├── modDateCalc.bas          ' 日期计算
│       ├── modUnitConvert.bas       ' 单位转换
│       ├── modSpecialChars.bas      ' 特殊字符扫描
│       └── modFileListing.bas       ' 文件目录生成
├── tests/                       # [新建] 测试资源
│   ├── test_vba_workbook.xlsm   #   含宏测试工作簿（内置测试数据和测试运行器）
│   └── test_cdm_plus.py         #   cdm_toolkit_plus.py 的Python测试
└── reference/                   # [已有] 参考材料
```

---

## 3. VBA 编码规范

所有 VBA 模块统一遵循以下约定：

| 规范项 | 规则 |
|--------|------|
| **命名** | 公开过程使用 `DM_<功能名>` 前缀，PascalCase；私有过程使用 camelCase |
| **参数化** | 所有可配置值均通过参数暴露，无硬编码范围/值 |
| **错误处理** | 每个公开的 Sub/Function 均有 `On Error GoTo` 并给出明确的错误提示 |
| **注释** | 公开接口使用中英双语注释；内部逻辑使用英文注释 |
| **兼容性** | Excel 2010+；64位安全 |
| **依赖** | 仅使用标准 VBA 引用，使用后期绑定避免外部 COM 库依赖 |
| **模块头** | 每个模块开头包含：名称、描述、作者、版本、依赖列表 |

---

## 4. 模块详细规格

### 4.1 modTextStandardise — 文本术语标准化

**来源**: Word文档 "筛选重复"（Cells.Replace部分） + Python `standardise_terms()`

**公开接口**:
```vb
Public Sub DM_TextStandardise(ByVal ws As Worksheet, _
    Optional ByVal termMap As Scripting.Dictionary)
' 使用CDM标准术语映射表，在整个工作表中批量替换非标准术语。
' 内置映射表：Recovered→Recovered/Resolved、Ongoing→Not Recovered/Not Resolved 等。

Public Sub DM_TextReplace(ByVal ws As Worksheet, _
    ByVal oldText As String, ByVal newText As String, _
    Optional ByVal lookAt As XlLookAt = xlPart)
' 单术语替换便捷方法。
```

**内置术语映射表**（与 Python `TERM_MAP` 对齐）：
- "Recovered" → "Recovered / Resolved"
- "Ongoing" → "Not Recovered / Not Resolved"
- "Unknown" → "Unknown"
- "Reasonable Possib" → "Reasonable Possibility"
- "Not Recovered/Not Resolved" → "Not Recovered / Not Resolved"
- "Recovered/Resolved" → "Recovered / Resolved"
- "Recovered/Resolved with Sequelae" → "Recovered / Resolved with Sequelae"

**相比原始VBA的改进**：
- 参数化（原始代码硬编码了 Replace 调用）
- 术语映射表存储在 Dictionary 中，而非内联代码
- 返回替换次数统计

---

### 4.2 modHighlightDup — 重复值高亮

**来源**: Word文档 "筛选重复"（高亮部分） + Python `highlight_duplicates()`

**公开接口**:
```vb
Public Sub DM_HighlightDuplicates(ByVal ws As Worksheet, _
    Optional ByVal colStart As Long = 7, _
    Optional ByVal colEnd As Long = 11, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal highlightColor As Long = 65535)
' 在指定列范围内高亮连续重复的单元格。
' 同时将当前行和上一行标记为黄色。

Public Sub DM_HighlightDupSingleCol(ByVal ws As Worksheet, _
    Optional ByVal col As Long = 5)
' 便捷方法：单列重复高亮。
```

**相比原始VBA的改进**：
- 列范围参数化（原始：硬编码 7-11）
- 颜色参数化（原始：硬编码 65535）
- 避免使用 `.Select`（原始使用 `Cells(i,X).Select`）

---

### 4.3 modBlankRows — 空白行管理

**来源**: Word文档 `Insert5` + `Insertb` + Python `insert_blank_every_n()` / `insert_blank_between_groups()` / `delete_blank_rows()`

**公开接口**:
```vb
Public Sub DM_InsertBlankEveryN(ByVal ws As Worksheet, _
    Optional ByVal n As Long = 5, _
    Optional ByVal colLetter As String = "A")
' 每N行插入一行空白。自底向上执行，保持行号正确。

Public Sub DM_InsertBlankBetweenGroups(ByVal ws As Worksheet, _
    Optional ByVal colLetter As String = "A")
' 当指定列的值发生变化时，在组间插入空白行。

Public Function DM_DeleteBlankRows(ByVal ws As Worksheet) As Long
' 删除完全空白的行。返回删除行数。
```

**相比原始VBA的改进**：
- `DM_DeleteBlankRows` 为新增功能（来自Python，原始VBA无此功能）
- 合并为一个模块（原始分散在 `Insert5` 和 `Insertb` 两个独立宏中）
- `colLetter` 参数替代硬编码的 "A" 列

---

### 4.4 modTOC — 目录生成器

**来源**: Word文档 `TOC` + `DeriveLink` + Python `generate_toc()` / `generate_sas_derive_toc()`

**公开接口**:
```vb
Public Sub DM_GenerateTOC(ByVal wb As Workbook, _
    Optional ByVal tocName As String = "TOC", _
    Optional ByVal excludeSheets As String = "")
' 创建/重建目录sheet，包含指向所有可见sheet的超链接。
' 在每个被链接的sheet上添加返回箭头形状。

Public Sub DM_GenerateDeriveTOC(ByVal wb As Workbook, _
    Optional ByVal descCell As String = "V2", _
    Optional ByVal tocName As String = "TOC")
' SAS Derive风格目录：A列=超链接sheet名称，B列=从descCell提取的描述。
```

**相比原始VBA的改进**：
- 合并了 `TOC` 和 `DeriveLink`（原始代码90%重复）
- 新增 `excludeSheets` 参数（逗号分隔的排除列表）
- 清理了错误处理（原始使用笼统的 `On Error Resume Next`）

---

### 4.5 modSheetSplit — 工作表拆分

**来源**: Word文档 `工作表拆分2` + Python `split_sheet_by_column()`

**公开接口**:
```vb
Public Sub DM_SplitSheetByColumn(ByVal ws As Worksheet, _
    Optional ByVal splitCol As String = "A", _
    Optional ByVal headerRows As Long = 1)
' 按指定列的唯一值将工作表拆分为多个sheet。

Public Sub DM_SplitSheetWithProgress(ByVal ws As Worksheet, _
    Optional ByVal splitCol As String = "A", _
    Optional ByVal headerRows As Long = 1)
' 同上，但在状态栏显示进度指示器，适用于大数据量。
```

**相比原始VBA的改进**：
- 使用 Dictionary 去重（原始使用 Collection，速度较慢）
- 创建sheet前检测名称冲突
- 新增 `DM_SplitSheetWithProgress` 带 `Application.StatusBar` 进度更新

---

### 4.6 modWorkbookMerge — 工作簿合并

**来源**: Word文档 `Books2Sheets0xls` / `Books2Sheets1xlsx` / `Books2Sheets2` + Python `merge_workbooks()`

**公开接口**:
```vb
Public Sub DM_MergeWorkbooksDialog(Optional ByVal createTOC As Boolean = True)
' 文件对话框选择 → 将选中的工作簿合并到新工作簿。

Public Sub DM_MergeWorkbooksFolder(ByVal folderPath As String, _
    Optional ByVal filePattern As String = "*.xlsx", _
    Optional ByVal createTOC As Boolean = True)
' 合并指定文件夹中所有匹配的文件。

Public Sub DM_MergeWorkbooksList(ByVal filePaths As Collection, _
    Optional ByVal outputPath As String = "", _
    Optional ByVal createTOC As Boolean = True)
' 从明确的文件路径列表合并。
```

**相比原始VBA的改进**：
- 三个原始宏（`Books2Sheets0xls`、`Books2Sheets1xlsx`、`Books2Sheets2`）合并为一个模块
- Sheet名称冲突解决（添加计数器后缀）
- 同时支持 `.xls` 和 `.xlsx`
- 合并每个源工作簿的所有sheet（原始 `Books2Sheets2` 仅复制第一个sheet）

---

### 4.7 modSheetManager — 工作表管理

**来源**: DMⅡ常用宏.xlsx 需求 #1, #2, #7（无原始VBA源码） + Python `batch_rename_sheets()` / `list_sheet_names()` / `delete_hidden_sheets()`

**公开接口**:
```vb
Public Function DM_ListSheetNames(ByVal wb As Workbook, _
    Optional ByVal includeHidden As Boolean = False) As String()
' 返回sheet名称数组。

Public Sub DM_BatchRenameSheets(ByVal wb As Workbook, _
    ByVal nameMap As Scripting.Dictionary)
' 批量重命名多个sheet。Key=旧名称，Value=新名称。
' Sheet名称自动截断到Excel的31字符限制。

Public Function DM_DeleteHiddenSheets(ByVal wb As Workbook) As String()
' 删除所有隐藏sheet。返回被删sheet的名称数组。

Public Sub DM_DeleteSheets(ByVal wb As Workbook, _
    ParamArray sheetNames() As Variant)
' 删除指定sheet（带确认提示）。
```

**说明**: 全新VBA模块 — 从 Python `cdm_toolkit.py` 反向编写。

---

### 4.8 modTrackChanges — 修改痕迹跟踪

**来源**: DMⅡ常用宏.xlsx 需求 #8（无原始VBA源码） + Python `track_changes_add_comment()`

**公开接口**:
```vb
Public Sub DM_TrackChanges(ByVal ws As Worksheet, _
    ByVal originalFilePath As String, _
    Optional ByVal addComments As Boolean = True, _
    Optional ByVal highlightChanges As Boolean = True)
' 将当前工作表与原始备份副本进行比较。
' 修改过的单元格：绿色填充 + 红色边框 + 批注 "旧值: X → 新值: Y"。

Public Sub DM_ClearTrackChanges(ByVal ws As Worksheet)
' 清除所有修改痕迹标记（批注、填充、边框）。

Public Sub DM_ChangesReport(ByVal ws As Worksheet, _
    ByVal originalFilePath As String)
' 生成新的"修改报告"sheet，包含结构化差异日志：
'   A列: 单元格地址, B列: 变更类型 (新增/删除/修改)
'   C列: 旧值, D列: 新值
```

**关键改进**: Python版本仅用颜色标记（openpyxl批注支持有限）。VBA版本可添加真正的Excel批注，显示旧值→新值，满足原始需求。

---

### 4.9 modDateCalc — 日期计算器

**来源**: DMⅡ常用宏.xlsx 需求 #9, #10（无原始VBA源码） + Python `add_days_to_column()` / `calc_date_diff()`

**公开接口**:
```vb
Public Sub DM_AddDaysToColumn(ByVal ws As Worksheet, _
    ByVal colLetter As String, _
    ByVal days As Long, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal dateFormat As String = "YYYY-MM-DD")
' 对指定列中的每个日期值加/减天数。

Public Sub DM_CalcDateDiff(ByVal ws As Worksheet, _
    ByVal colStart As String, _
    ByVal colEnd As String, _
    ByVal resultCol As String, _
    Optional ByVal rowStart As Long = 2)
' 计算 (结束日期 - 开始日期) 的天数差，写入结果列。

Public Sub DM_DateAddSelection(ByVal days As Long)
' 便捷方法：对当前选中的所有日期单元格加天数。
```

**说明**: 全新VBA模块 — 从 Python 反向编写。

---

### 4.10 modUnitConvert — 实验室单位转换

**来源**: DMⅡ常用宏.xlsx 需求 #11（无原始VBA源码） + Python `convert_lab_units()`

**公开接口**:
```vb
Public Sub DM_ConvertLabUnits(ByVal ws As Worksheet, _
    ByVal testCol As String, _        ' 检验项目列
    ByVal resultCol As String, _      ' 结果值列
    ByVal unitCol As String, _        ' 单位列
    ByVal standardUnit As String, _   ' 目标标准单位
    Optional ByVal testFilter As String = "")
' 使用内置转换因子将实验室结果转换为标准单位。
' 无法转换的行标记为红色。

Public Function DM_GetSupportedConversions() As String()
' 列出所有可用的单位转换。

Public Sub DM_AddUnitConversion(ByVal fromUnit As String, _
    ByVal toUnit As String, ByVal factor As Double)
' 在运行时扩展转换表。
```

**内置转换表**（与 Python `UNIT_CONVERSIONS` 对齐）：24对转换因子，覆盖质量（g/mg/ug/kg）、体积（L/mL/dL）、长度（cm/mm/m）、浓度（g/L↔mg/mL↔mg/dL↔umol/L）和时间（h/min/d）。

**说明**: 全新VBA模块 — 从 Python 反向编写。

---

### 4.11 modSpecialChars — 特殊字符扫描

**来源**: DMⅡ常用宏.xlsx 需求（Rave/Clinflash 特殊字符检查） + Python `scan_special_chars()`

**公开接口**:
```vb
Public Function DM_ScanSpecialChars(ByVal ws As Worksheet, _
    Optional ByVal allowedPattern As String = "") As Scripting.Dictionary
' 扫描工作表中包含允许字符集以外字符的单元格。
' 默认允许: [a-zA-Z0-9\s\-_.,:;()/]
' 返回 Dictionary(单元格地址 → 单元格值)。

Public Sub DM_HighlightSpecialChars(ByVal ws As Worksheet, _
    Optional ByVal highlightColor As Long = vbRed)
' 查找并高亮包含非ASCII/非标准字符的单元格。

Public Sub DM_EDCSpecialCharReport(ByVal ws As Worksheet, _
    Optional ByVal system As String = "Rave")
' 生成EDC专用报告（Rave 或 Clinflash）。
' 列出所有特殊字符及其单元格地址，用于提交Query。
```

**说明**: 全新VBA模块，包含EDC系统（Rave/Clinflash）专用报告功能，超出Python当前功能范围。

---

### 4.12 modFileListing — 文件目录生成

**来源**: Word文档 `autogetFld` + `Searchfiletohyperlinks` + `ListAllFso`  
**Python对应**: 新建 — `cdm_toolkit_plus.py` > `generate_file_listing()`

**公开接口**:
```vb
Public Sub DM_GenerateFileListing(ByVal ws As Worksheet, _
    Optional ByVal folderPath As String = "", _
    Optional ByVal includeSubFolders As Boolean = True, _
    Optional ByVal fileFilter As String = "*.*")
' 列出文件夹中所有文件并添加超链接。
' A列: 文件夹路径, B列: 文件名（带超链接）。

Public Sub DM_GenerateFolderTree(ByVal ws As Worksheet, _
    Optional ByVal folderPath As String = "")
' 递归树形视图展示，带缩进。
```

**相比原始VBA的改进**：
- 合并了 `autogetFld`/`Searchfiletohyperlinks`/`ListFilesTest`/`ListAllFso`（4个过程 → 2个）
- 统一树形格式，缩进一致
- 新增 `fileFilter` 参数用于文件类型过滤

---

## 5. Python 补充: cdm_toolkit_plus.py

### 新增函数

| 函数 | 描述 | 对应VBA模块 |
|------|------|-------------|
| `generate_file_listing(folder_path, output_path)` | 生成带超链接的文件夹内容Excel | modFileListing |
| `generate_folder_tree(folder_path, output_path)` | 递归文件夹树形结构Excel | modFileListing |
| `track_changes_with_comments(wb, orig_path)` | 使用COM添加真正的Excel批注来跟踪修改 | modTrackChanges |
| `export_vba_module(xlsm_path, bas_path)` | 将.bas模块嵌入.xlsm工作簿 | (安装辅助) |

### API 设计

```python
def generate_file_listing(folder_path: str, output_path: str,
                          include_subfolders: bool = True,
                          file_pattern: str = "*.*") -> Workbook:
    """生成带超链接的文件夹内容Excel。返回Workbook对象。"""

def generate_folder_tree(folder_path: str, output_path: str) -> Workbook:
    """生成递归树形结构的文件夹展示Excel。"""

def track_changes_with_comments(wb: Workbook, original_path: str) -> int:
    """使用COM添加真正的Excel批注来跟踪修改。
    如果COM不可用，退回到颜色标记方式。返回变更数量。"""

def export_vba_module(xlsm_path: str, bas_path: str) -> bool:
    """通过COM将.bas模块嵌入.xlsm工作簿。"""
```

---

## 6. 测试策略

### VBA 测试

- `test_vba_workbook.xlsm`：含宏测试工作簿，包含：
  - 预加载的CDM模拟数据（AE、LB、DM sheet）
  - 测试运行器 Sub，依次调用每个宏并验证输出
  - 通过注释记录预期结果

### Python 测试

- `test_cdm_plus.py`：延续已有 `test_cdm_toolkit.py` 的测试模式
  - 测试 `cdm_toolkit_plus.py` 中所有新增函数
  - 集成测试：Python ↔ VBA 往返验证

---

## 7. 实现顺序

1. **基础模块**（先易后难，建立规范参考）: modTextStandardise、modHighlightDup、modBlankRows
2. **核心模块**（中等复杂度，有原始VBA参考）: modTOC、modSheetSplit、modWorkbookMerge
3. **全新编写**（从Python反向编写，纯新代码）: modSheetManager、modDateCalc、modUnitConvert
4. **高级模块**（逻辑更复杂）: modTrackChanges、modSpecialChars
5. **收尾**（文件目录 + Python补充）: modFileListing + cdm_toolkit_plus.py
6. **集成**（主模块 + 安装 + 文档 + 测试）: CDM_Toolkit.bas、CDM_Install.bas、README.md、测试

---

## 8. 风险与缓解

| 风险 | 缓解措施 |
|------|----------|
| VBA Dictionary 需要 `Microsoft Scripting Runtime` 引用 | 使用后期绑定（`CreateObject("Scripting.Dictionary")`）避免引用依赖 |
| Excel 31字符sheet名称限制（拆分/合并时） | 截断 + 计数器后缀逻辑 |
| 大数据量导致VBA性能问题 | `Application.ScreenUpdating=False`、`Application.Calculation=xlCalculationManual` |
| .bas 文件中文字符编码问题 | 以 UTF-8 BOM 格式保存 .bas；在中英文Excel环境中分别测试导入 |
