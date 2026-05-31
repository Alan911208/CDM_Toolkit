# sas2xls

## 中文说明

### 功能
通过 DDE（动态数据交换）将 SAS 数据集导出到指定的 Excel 工作表。这是传统的导出方式，需要 Windows 环境且 Excel 必须正在运行。

**强烈建议**：在现代 SAS 环境中（9.4 及以上版本），请使用 sas2xlsx（宏 13），它使用 PROC EXPORT dbms=xlsx，无需 Excel 运行且更加稳定可靠。

### 前置条件
- Windows 操作系统
- Microsoft Excel 已安装并在调用宏时保持运行
- SAS 会话支持 DDE

### 参数
| 参数       | 说明                                    | 默认值    |
|------------|-----------------------------------------|-----------|
| `inset`    | 输入的 SAS 数据集（一级或二级名称）     | (必填)    |
| `outsheet` | 目标 Excel 工作表名称                    | Sheet1    |
| `outfile`  | 目标 Excel 工作簿的完整路径（必填）     | (必填)    |

### 示例
```sas
/* 将 work.AE 导出到 Excel */
%sas2xls(inset=AE, outsheet=AE_Data, outfile=C:\output\ae_report.xls);

/* 使用二级数据集名 */
%sas2xls(inset=myLib.demog, outsheet=Demographics, outfile=C:\output\study.xls);
```

### 注意事项
- 并非所有 SAS 安装都支持 DDE（如 SAS/Connect、SAS Server 环境）
- 必须在调用此宏之前手动打开 Excel，或通过 `X` 命令启动 Excel
- DDE 三元组格式可能因 SAS/Excel 版本而异
- 本宏不处理高级格式（字体、颜色、合并单元格等）
- 本宏输出详细注释模板，请根据实际环境取消注释并调整

---

## English

### Purpose
Export a SAS dataset to a specific Excel sheet using DDE (Dynamic Data Exchange). This is a legacy approach requiring Windows and a running Excel instance.

**Strongly recommended**: For modern SAS (9.4+), use sas2xlsx (macro 13) which uses PROC EXPORT dbms=xlsx — no running Excel required, more stable and reliable.

### Prerequisites
- Windows operating system
- Microsoft Excel installed and running at time of macro execution
- SAS session with DDE support

### Parameters
| Parameter  | Description                                        | Default    |
|------------|----------------------------------------------------|------------|
| `inset`    | Input SAS dataset (one-level or two-level name)    | (required) |
| `outsheet` | Target Excel sheet name                            | Sheet1     |
| `outfile`  | Full path to the target Excel workbook (required)  | (required) |

### Example
```sas
/* Export work.AE to Excel */
%sas2xls(inset=AE, outsheet=AE_Data, outfile=C:\output\ae_report.xls);

/* Use two-level dataset name */
%sas2xls(inset=myLib.demog, outsheet=Demographics, outfile=C:\output\study.xls);
```

### Important Notes
- Not all SAS installations support DDE (e.g., SAS/Connect, SAS Server environments)
- Excel must be opened manually before calling this macro, or launched via the `X` command
- DDE triplet format may differ across SAS/Excel versions
- This macro does not handle advanced formatting (fonts, colors, merged cells, etc.)
- The macro outputs a detailed commented template; uncomment and adjust for your environment
