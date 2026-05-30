# xls2sas

## 中文说明

### 功能
将 Excel 工作簿中指定的工作表导入为 SAS 数据集。此宏是对 PROC IMPORT 的便捷封装，提供了参数验证、文件存在性检查和合理的默认值。

### 参数
| 参数       | 说明                                                    | 默认值  |
|------------|---------------------------------------------------------|---------|
| `file`     | Excel 工作簿的完整路径（必填）                          | (必填)  |
| `insheet`  | 要导入的工作表名称（必填）                              | (必填)  |
| `outset`   | 输出 SAS 数据集名称（必填）                             | (必填)  |
| `outlib`   | 输出逻辑库名称                                          | work    |
| `dbms`     | 文件类型：xlsx, xls 或 excel                            | xlsx    |
| `getnames` | 是否将第一行作为变量名：YES 或 NO                       | YES     |
| `mixed`    | 是否允许列中混合字符和数值：YES 或 NO                   | NO      |
| `usedate`  | 是否对日期列使用 SAS 日期格式：YES 或 NO                | YES     |
| `scantext` | 是否扫描文本列以确定最大长度：YES 或 NO                 | YES     |

### 依赖
- 需要 SAS/ACCESS to PC Files 许可

### 示例
```sas
/* 基本用法：导入指定的工作表 */
%xls2sas(file=C:\data\study.xlsx, insheet=Demog, outset=demog);

/* 指定输出库和 dbms 类型 */
%xls2sas(file=C:\data\lab.xlsx, insheet=Chemistry, outset=lb,
         outlib=mylib, dbms=xls);

/* 第一行不是表头 */
%xls2sas(file=C:\data\raw.xlsx, insheet=Data, outset=raw,
         getnames=NO);
```

### 注意事项
- 在导入前会检查 Excel 文件是否存在
- 如果输出数据集已存在，会自动删除后重建
- PROC IMPORT 会根据数据值自动推断变量类型（字符/数值）
- 大数据量导入时可能较慢，建议预检数据行数

---

## English

### Purpose
Import a specified sheet from an Excel workbook into a SAS dataset. A convenience wrapper around PROC IMPORT with parameter validation, file existence checking, and sensible defaults.

### Parameters
| Parameter  | Description                                               | Default   |
|------------|-----------------------------------------------------------|-----------|
| `file`     | Full path to the Excel workbook (required)                | (required)|
| `insheet`  | Sheet name to import (required)                            | (required)|
| `outset`   | Output SAS dataset name (required)                         | (required)|
| `outlib`   | Output library name                                        | work      |
| `dbms`     | File type: xlsx, xls, or excel                             | xlsx      |
| `getnames` | Treat first row as variable names: YES or NO               | YES       |
| `mixed`    | Allow mixed char/num in a column: YES or NO                | NO        |
| `usedate`  | Use SAS date format for date columns: YES or NO            | YES       |
| `scantext` | Scan text columns for max length: YES or NO                | YES       |

### Dependencies
- Requires SAS/ACCESS to PC Files license

### Example
```sas
/* Basic usage: import a specific sheet */
%xls2sas(file=C:\data\study.xlsx, insheet=Demog, outset=demog);

/* Specify output library and file type */
%xls2sas(file=C:\data\lab.xlsx, insheet=Chemistry, outset=lb,
         outlib=mylib, dbms=xls);

/* First row is NOT headers */
%xls2sas(file=C:\data\raw.xlsx, insheet=Data, outset=raw,
         getnames=NO);
```

### Notes
- Checks if the Excel file exists before attempting import
- If the output dataset already exists, it is automatically dropped and recreated
- PROC IMPORT auto-infers variable types (character/numeric) from data values
- Large imports may be slow; consider pre-checking row counts
