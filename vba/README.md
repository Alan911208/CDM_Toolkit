# CDM Excel VBA Toolkit v1.0

临床数据管理 (Clinical Data Management) Excel VBA 宏工具包。

## 快速开始

### 安装

**方法一：手动导入（推荐）**
1. 打开 Excel，按 `Alt+F11` 进入 VBA 编辑器
2. `File` → `Import File` → 选择 `vba/modules/` 下所有 `.bas` 文件
3. 再导入 `vba/CDM_Toolkit.bas`

**方法二：一键安装到 Personal.xlsb**
1. 先导入所有 `vba/modules/*.bas` + `vba/CDM_Install.bas` 到任意工作簿
2. 运行 `DM_InstallAll` 宏
3. 所有宏将被复制到 PERSONAL.XLSB，之后任意 Excel 文件都可用

### 使用

在Excel中按 `Alt+F8`，选择要运行的宏，点击"执行"。

---

## 模块列表

### 数据清理

| 宏名 | 功能 | 用法示例 |
|------|------|---------|
| `DM_TextStandardise` | 批量标准化CDM术语 | `DM_TextStandardise ActiveSheet` |
| `DM_HighlightDuplicates` | 高亮连续重复的单元格（黄色） | `DM_HighlightDuplicates ActiveSheet` |
| `DM_InsertBlankEveryN` | 每N行插入空白行 | `DM_InsertBlankEveryN ActiveSheet, 5` |
| `DM_InsertBlankBetweenGroups` | 组间插入空白行 | `DM_InsertBlankBetweenGroups ActiveSheet, "A"` |
| `DM_DeleteBlankRows` | 删除完全空白的行 | `DM_DeleteBlankRows ActiveSheet` |

### 文档管理

| 宏名 | 功能 | 用法示例 |
|------|------|---------|
| `DM_GenerateTOC` | 生成带超链接的目录sheet | `DM_GenerateTOC ActiveWorkbook` |
| `DM_GenerateDeriveTOC` | SAS Derive风格目录 | `DM_GenerateDeriveTOC ActiveWorkbook, "V2"` |
| `DM_SplitSheetByColumn` | 按列值拆分sheet | `DM_SplitSheetByColumn ActiveSheet, "A"` |
| `DM_MergeWorkbooksDialog` | 合并多个工作簿 | `DM_MergeWorkbooksDialog` |

### 工作表管理

| 宏名 | 功能 | 用法示例 |
|------|------|---------|
| `DM_ListSheetNames` | 获取所有sheet名称 | `DM_ListSheetNames ActiveWorkbook` |
| `DM_BatchRenameSheets` | 批量重命名sheet | `DM_BatchRenameSheets ActiveWorkbook, nameMap` |
| `DM_DeleteHiddenSheets` | 删除所有隐藏sheet | `DM_DeleteHiddenSheets ActiveWorkbook` |

### 审核追溯

| 宏名 | 功能 | 用法示例 |
|------|------|---------|
| `DM_TrackChanges` | 比较原版并标记修改痕迹 | `DM_TrackChanges ActiveSheet, "C:\backup\orig.xlsx"` |
| `DM_ChangesReport` | 生成结构化修改报告 | `DM_ChangesReport ActiveSheet, "C:\backup\orig.xlsx"` |
| `DM_ScanSpecialChars` | 扫描特殊字符 | `DM_ScanSpecialChars ActiveSheet` |
| `DM_EDCSpecialCharReport` | 生成EDC特殊字符报告 | `DM_EDCSpecialCharReport ActiveSheet, "Rave"` |

### 计算转换

| 宏名 | 功能 | 用法示例 |
|------|------|---------|
| `DM_AddDaysToColumn` | 日期列加/减天数 | `DM_AddDaysToColumn ActiveSheet, "F", 30` |
| `DM_CalcDateDiff` | 计算两列日期天数差 | `DM_CalcDateDiff ActiveSheet, "F", "G", "M"` |
| `DM_ConvertLabUnits` | 实验室单位统一转换 | `DM_ConvertLabUnits ActiveSheet, "D", "E", "F", "mg/dL", "Glucose"` |

### 工具

| 宏名 | 功能 | 用法示例 |
|------|------|---------|
| `DM_GenerateFileListing` | 生成文件夹文件目录 | `DM_GenerateFileListing ActiveSheet` |
| `DM_GenerateFolderTree` | 生成文件夹树形结构 | `DM_GenerateFolderTree ActiveSheet` |

### 快捷入口

| 宏名 | 功能 |
|------|------|
| `DM_QuickClean` | 一键清理：术语标准化 + 重复高亮 + 删除空白行 |
| `DM_QuickDoc` | 一键文档：生成TOC目录 |
| `DM_About` | 显示版本信息 |

---

## 系统要求

- Excel 2010 或更高版本
- 无需额外 COM 引用（使用后期绑定）
- 安装到 Personal.xlsb 需要启用"信任对VBA工程对象的模型访问"
