# csv2sas

## 中文说明

### 功能
将 CSV 文件导入为 SAS 数据集，支持智能变量命名。采用三遍扫描方式：第一遍解析表头并清理变量名，第二遍检测每列的最大数据宽度，第三遍以最优长度读取数据。

### 参数
| 参数     | 说明                                                      | 默认值    |
|----------|-----------------------------------------------------------|-----------|
| `file`   | CSV 文件的完整路径（必填）                                | (必填)    |
| `dsn`    | 输出数据集名称                                            | csv2sas   |
| `naming` | 变量命名规则：`XLS2SAS` 或 `LABEL`                        | XLS2SAS   |
| `lrecl`  | INFILE 语句的逻辑记录长度                                  | 5000      |

### 命名规则说明
- **XLS2SAS**：变量命名为 `_c1_`, `_c2_`, ...，原始表头作为标签（label）
- **LABEL**：变量名由表头清理后生成（非法字符替换/删除），原始表头作为标签

### 示例
```sas
%csv2sas(file=C:\data\subject.csv, dsn=work.subj, naming=LABEL);
%csv2sas(file=C:\data\ae.csv, dsn=work.ae);
```

---

## English

### Purpose
Import a CSV file into a SAS dataset with intelligent variable naming. Uses a 3-pass import strategy: (1) parse headers and sanitize variable names, (2) detect maximum column widths, (3) read data with optimal string lengths.

### Parameters
| Parameter | Description                                                | Default   |
|-----------|------------------------------------------------------------|-----------|
| `file`    | Full path to the CSV file (required)                       | (required)|
| `dsn`     | Output dataset name                                        | csv2sas   |
| `naming`  | Variable naming convention: `XLS2SAS` or `LABEL`           | XLS2SAS   |
| `lrecl`   | Logical record length for INFILE statement                 | 5000      |

### Naming Convention
- **XLS2SAS**: Variables named `_c1_`, `_c2_`, ..., labels set from original headers
- **LABEL**: Variable names derived from cleaned headers (special chars replaced), labels from headers

### Example
```sas
%csv2sas(file=C:\data\subject.csv, dsn=work.subj, naming=LABEL);
%csv2sas(file=C:\data\ae.csv, dsn=work.ae);
```

### Notes
- Uses an internal `_dropIfExists` helper macro — included inline, no external dependency
- Trailing blanks in variable names stored as labels are preserved via `$` input format
- Variables whose cleaned name becomes empty or invalid are assigned `_col0`, `_col1`, etc.
