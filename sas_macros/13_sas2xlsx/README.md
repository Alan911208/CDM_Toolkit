# sas2xlsx

## 中文说明

### 功能
将 SAS 逻辑库中的所有数据集批量导出为一个 Excel 工作簿（.xlsx），每个数据集对应一个工作表。使用 PROC EXPORT 和 `dbms=xlsx` 引擎。

### 参数
| 参数   | 说明                                            | 默认值  |
|--------|-------------------------------------------------|---------|
| `lib`  | 要导出的 SAS 逻辑库名称（必填）                 | (必填)  |
| `dir`  | 输出目录路径（必填）                            | (必填)  |
| `name` | 输出工作簿文件名（不含扩展名，必填）            | (必填)  |
| `ext`  | 文件扩展名，通常为 xlsx（默认）或 xls           | xlsx    |

### 依赖
- 需要 SAS/ACCESS to PC Files 许可

### 示例
```sas
/* 将 WORK 库中的所有数据集导出为一个 Excel 文件 */
%sas2xlsx(lib=work, dir=C:\output, name=study_data);

/* 指定扩展名为 xls */
%sas2xlsx(lib=sashelp, dir=C:\output, name=class_data, ext=xls);
```

### 注意事项
- 导出的 Excel 文件中每个工作表名与 SAS 数据集名一致
- 若逻辑库中没有数据集，会输出 WARNING 并退出
- `label` 选项表示使用变量标签作为 Excel 列标题

---

## English

### Purpose
Batch-export all SAS datasets in a library to a single Excel workbook (.xlsx), one sheet per dataset. Uses PROC EXPORT with `dbms=xlsx`.

### Parameters
| Parameter | Description                                              | Default    |
|-----------|----------------------------------------------------------|------------|
| `lib`     | SAS library name to export from (required)               | (required) |
| `dir`     | Output directory path (required)                         | (required) |
| `name`    | Output workbook filename without extension (required)    | (required) |
| `ext`     | File extension, typically xlsx (default) or xls          | xlsx       |

### Dependencies
- Requires SAS/ACCESS to PC Files license

### Example
```sas
/* Export all WORK datasets to a single Excel file */
%sas2xlsx(lib=work, dir=C:\output, name=study_data);

/* Specify .xls extension */
%sas2xlsx(lib=sashelp, dir=C:\output, name=class_data, ext=xls);
```

### Notes
- Each sheet in the output workbook is named after the source SAS dataset
- If no datasets are found in the library, a WARNING is issued and the macro exits
- The `label` option exports variable labels as Excel column headers
