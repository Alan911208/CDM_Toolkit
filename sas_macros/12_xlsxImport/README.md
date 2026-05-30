# xlsxImport

## 中文说明

### 功能
将一个 Excel 工作簿中的所有工作表（sheet）批量导入为独立的 SAS 数据集。首先通过 LIBNAME 引擎读取工作表列表，然后对每个工作表调用 PROC IMPORT 进行导入。

### 参数
| 参数   | 说明                                | 默认值  |
|--------|-------------------------------------|---------|
| `file` | Excel 工作簿的完整路径 (.xlsx)      | (必填)  |
| `out`  | 目标 SAS 逻辑库名称                 | (必填)  |

### 依赖
- 需要 SAS/ACCESS to PC Files 许可（支持 LIBNAME XLSX 引擎和 PROC IMPORT dbms=xlsx）

### 示例
```sas
/* 初始化目标库 */
libname myout "C:\sasdata\imported";

/* 导入所有工作表 */
%xlsxImport(file=C:\data\study_data.xlsx, out=myout);

/* 查看导入的数据集 */
proc datasets lib=myout; quit;
```

### 注意事项
- 暂存表 `work._excel` 会在导入完成后自动删除
- 工作表名称中的 `$` 符号会被去除（`compress(memname, '$')`）
- 每个工作表的第一行会被视为变量名（`getnames=yes`）

---

## English

### Purpose
Batch-import all sheets from an Excel workbook into individual SAS datasets. Reads the sheet list via the LIBNAME engine, then calls PROC IMPORT for each sheet.

### Parameters
| Parameter | Description                                | Default    |
|-----------|--------------------------------------------|------------|
| `file`    | Full path to the Excel workbook (.xlsx)    | (required) |
| `out`     | Target SAS library name                    | (required) |

### Dependencies
- Requires SAS/ACCESS to PC Files license (for LIBNAME XLSX engine and PROC IMPORT dbms=xlsx)

### Example
```sas
/* Initialize target library */
libname myout "C:\sasdata\imported";

/* Import all sheets */
%xlsxImport(file=C:\data\study_data.xlsx, out=myout);

/* View imported datasets */
proc datasets lib=myout; quit;
```

### Notes
- Temporary table `work._excel` is dropped automatically after import
- The `$` suffix on sheet names is stripped via `compress(memname, '$')`
- The first row of each sheet is treated as variable names (`getnames=yes`)
