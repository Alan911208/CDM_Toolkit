# %textSplit

## Category
Data Manipulation & Cleaning

## Purpose / 目的

**English:** Split a long character variable (exceeding 200 characters) into multiple columns, each holding up to 200 characters. The original variable is dropped and replaced by numbered columns (`VAR1`, `VAR2`, ..., `VARN`). This macro uses multi-byte-aware functions (`ksubstr`, `klength`, `ktrim`) to correctly handle Unicode/CJK text without corrupting multi-byte character boundaries.

**中文：** 将超过 200 个字符的长字符变量拆分为多个列，每列最多 200 个字符。原始变量被删除并由编号列（`VAR1`、`VAR2`、...、`VARN`）替换。该宏使用多字节感知函数（`ksubstr`、`klength`、`ktrim`），确保正确处理 Unicode/CJK 文本，避免截断多字节字符边界。

## Parameters / 参数

| Parameter | Required | Description |
|-----------|----------|-------------|
| `dsn`     | Yes      | Dataset name to operate on (modified in place) |
| `var`     | Yes      | Character variable name to split |

## Output / 输出

The input dataset `&dsn` is modified in place:
- Original variable `&var` is dropped
- New variables `&var.1`, `&var.2`, ..., `&var.&_max` created
- Each new variable has length 200 and inherits the original variable's label suffixed with the segment number

## Usage Examples / 使用示例

### Example 1: Split a long narrative text field
```sas
data work.narr;
    length TEXT $600;
    TEXT = "This is a very long narrative... (500+ chars)";
run;
%textSplit(dsn=work.narr, var=TEXT);
/* Result: TEXT1, TEXT2, TEXT3 created, TEXT dropped */
*/

### Example 2: Split Chinese text (multi-byte characters)
```sas
data work.comments;
    length COMMENT $1000;
    COMMENT = "这是一段非常长的中文注释文本...";
run;
%textSplit(dsn=work.comments, var=COMMENT);
/* Multi-byte-safe: COMMENT1, COMMENT2, etc. */
```

### Example 3: Conditional usage — only split if length exceeded
```sas
proc sql noprint;
    select max(lengthn(var)) into :_len from work.data;
quit;
%if &_len > 200 %then %do;
    %textSplit(dsn=work.data, var=var);
%end;
```

## Notes / 注意事项

- The macro determines the number of segments automatically by finding the longest value of `&var` across all rows (using `max(_n)`).
- If `&var` is empty in all rows, the macro prints a WARNING and exits without modifying the dataset.
- The original variable `&var` is DROPPED from the dataset — ensure you do not need it afterward.
- Multi-byte character boundaries are respected: the macro uses `ksubstr`/`klength`/`ktrim` for CJK and Unicode safety.
- The macro creates a temporary work dataset `_cut1` which is dropped afterward.
- The internal helper `_dropIfExists` is included in this standalone .sas file.

## Source Reference / 来源

`DM_Toolbox.sas` — Tool 5, Category 2: Data Manipulation & Cleaning
