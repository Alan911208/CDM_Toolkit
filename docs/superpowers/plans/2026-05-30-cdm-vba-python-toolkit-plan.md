# CDM VBA & Python 工具包 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 开发12个VBA模块 + 1个Python补充文件，实现CDM Excel工作流的完整工具包，与已有 cdm_toolkit.py 双向对齐。

**Architecture:** 12个独立 .bas 文件放在 `vba/modules/`，每个模块包含1-3个公开过程；1个主汇总模块 `vba/CDM_Toolkit.bas`；1个安装模块 `vba/CDM_Install.bas`；Python补充文件 `cdm_toolkit_plus.py` 放在项目根目录。

**Tech Stack:** VBA 7.1 (Excel 2010+), Python 3.10+ with openpyxl + xlwings, Scripting.Dictionary (后期绑定)

---

## 文件结构

```
Tools_DM/
├── cdm_toolkit_plus.py          # [新建] Python补充
├── vba/                          # [新建]
│   ├── modules/                  # [新建] 12个独立VBA模块
│   │   ├── modTextStandardise.bas
│   │   ├── modHighlightDup.bas
│   │   ├── modBlankRows.bas
│   │   ├── modTOC.bas
│   │   ├── modSheetSplit.bas
│   │   ├── modWorkbookMerge.bas
│   │   ├── modSheetManager.bas
│   │   ├── modTrackChanges.bas
│   │   ├── modDateCalc.bas
│   │   ├── modUnitConvert.bas
│   │   ├── modSpecialChars.bas
│   │   └── modFileListing.bas
│   ├── CDM_Toolkit.bas           # [新建] 主汇总模块
│   ├── CDM_Install.bas           # [新建] 一键安装模块
│   └── README.md                 # [新建] 使用说明书
└── tests/                        # [新建]
    ├── test_cdm_plus.py          # [新建] Python测试
    └── test_vba_workbook.bas     # [新建] VBA测试代码（导入xlsm用）
```

---

## 阶段一：基础模块（先易后难，建立规范参考）

### Task 1: modTextStandardise.bas — 文本术语标准化

**Files:**
- Create: `vba/modules/modTextStandardise.bas`

- [ ] **Step 1: 创建模块文件，包含完整代码**

```vb
Attribute VB_Name = "modTextStandardise"
Option Explicit

'============================================================================
' 模块: modTextStandardise
' 描述: CDM文本术语标准化 — 批量查找替换非标准术语为标准CDM术语
'        Text/Term Standardisation — batch replace non-standard terms with CDM standard terms
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无（使用后期绑定 Scripting.Dictionary）
'============================================================================

' --- 私有模块级变量 ---
Private mTermMap As Object ' Scripting.Dictionary

'============================================================================
' DM_InitTermMap
' 初始化内置CDM标准术语映射表（与 Python TERM_MAP 完全对齐）
'============================================================================
Private Sub DM_InitTermMap()
    If Not mTermMap Is Nothing Then Exit Sub
    
    Set mTermMap = CreateObject("Scripting.Dictionary")
    With mTermMap
        .Add "Recovered", "Recovered / Resolved"
        .Add "Ongoing", "Not Recovered / Not Resolved"
        .Add "Unknown", "Unknown"
        .Add "Reasonable Possib", "Reasonable Possibility"
        .Add "Not Recovered/Not Resolved", "Not Recovered / Not Resolved"
        .Add "Recovered/Resolved", "Recovered / Resolved"
        .Add "Recovered/Resolved with Sequelae", "Recovered / Resolved with Sequelae"
    End With
End Sub

'============================================================================
' DM_GetDefaultTermMap
' 获取默认术语映射表（供外部模块引用）
' 返回: Scripting.Dictionary
'============================================================================
Public Function DM_GetDefaultTermMap() As Object
    Call DM_InitTermMap
    Set DM_GetDefaultTermMap = mTermMap
End Function

'============================================================================
' DM_TextStandardise
' 使用CDM标准术语映射表批量替换整个工作表中的非标准术语
' 参数:
'   ws       — 目标工作表
'   termMap  — 可选，自定义术语映射表 {旧文本: 新文本}；省略则使用内置CDM映射表
' 返回: Long — 替换次数
' 用法:
'   Dim count As Long
'   count = DM_TextStandardise(ActiveSheet)
'   ' 或使用自定义映射表:
'   Dim customMap As Object: Set customMap = CreateObject("Scripting.Dictionary")
'   customMap.Add "AE", "Adverse Event"
'   count = DM_TextStandardise(ActiveSheet, customMap)
'============================================================================
Public Function DM_TextStandardise(ByVal ws As Worksheet, _
    Optional ByVal termMap As Object = Nothing) As Long
    
    On Error GoTo ErrHandler
    
    Dim cell As Range
    Dim key As Variant
    Dim count As Long
    Dim map As Object
    
    ' 确定使用的映射表
    If termMap Is Nothing Then
        Call DM_InitTermMap
        Set map = mTermMap
    Else
        Set map = termMap
    End If
    
    If map.Count = 0 Then
        DM_TextStandardise = 0
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    ' 遍历所有有值的单元格
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            For Each key In map.Keys
                If InStr(1, cell.Value, CStr(key), vbTextCompare) > 0 Then
                    cell.Value = Replace(cell.Value, CStr(key), CStr(map(key)), 1, -1, vbTextCompare)
                    count = count + 1
                End If
            Next key
        End If
    Next cell
    
    Application.ScreenUpdating = True
    DM_TextStandardise = count
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    DM_TextStandardise = -1
    MsgBox "DM_TextStandardise 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_TextReplace
' 单术语替换便捷方法
' 参数:
'   ws       — 目标工作表
'   oldText  — 要查找的文本
'   newText  — 替换为的文本
'   lookAt   — xlPart（部分匹配）或 xlWhole（完全匹配），默认 xlPart
' 用法:
'   DM_TextReplace ActiveSheet, "Recovered", "Recovered / Resolved"
'============================================================================
Public Sub DM_TextReplace(ByVal ws As Worksheet, _
    ByVal oldText As String, _
    ByVal newText As String, _
    Optional ByVal lookAt As XlLookAt = xlPart)
    
    On Error GoTo ErrHandler
    
    If Len(oldText) = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
    ws.Cells.Replace What:=oldText, Replacement:=newText, _
        LookAt:=lookAt, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_TextReplace 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证模块可导入VBA**

在Excel VBA编辑器中：File → Import File → 选择 `vba/modules/modTextStandardise.bas`，确认无编译错误。

---

### Task 2: modHighlightDup.bas — 重复值高亮

**Files:**
- Create: `vba/modules/modHighlightDup.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modHighlightDup"
Option Explicit

'============================================================================
' 模块: modHighlightDup
' 描述: CDM重复值高亮 — 标记相邻行中相同列值连续重复的单元格
'        Duplicate Highlighting — highlight consecutive duplicate cells across columns
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

'============================================================================
' DM_HighlightDuplicates
' 在指定列范围内高亮连续重复的单元格（当前行和上一行同时标记）
' 参数:
'   ws             — 目标工作表
'   colStart       — 起始列号（1-based），默认 7（G列）
'   colEnd         — 结束列号（1-based），默认 11（K列）
'   rowStart       — 数据起始行号，默认 2（第1行为表头）
'   highlightColor — 高亮颜色（VBA颜色常量），默认 vbYellow (65535)
' 用法:
'   DM_HighlightDuplicates ActiveSheet                    ' 默认 G-K列
'   DM_HighlightDuplicates ActiveSheet, 5, 5, 2, vbRed   ' 仅E列，红色
'============================================================================
Public Sub DM_HighlightDuplicates(ByVal ws As Worksheet, _
    Optional ByVal colStart As Long = 7, _
    Optional ByVal colEnd As Long = 11, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal highlightColor As Long = 65535)
    
    On Error GoTo ErrHandler
    
    Dim lastRow As Long
    Dim colIdx As Long
    Dim rowIdx As Long
    Dim currVal As Variant
    Dim prevVal As Variant
    
    ' 验证参数
    If colStart < 1 Or colEnd < colStart Or rowStart < 1 Then
        MsgBox "DM_HighlightDuplicates: 参数无效。" & vbCrLf & _
            "colStart=" & colStart & ", colEnd=" & colEnd & ", rowStart=" & rowStart, _
            vbExclamation, "CDM Toolkit"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colStart).End(xlUp).Row
    If lastRow <= rowStart Then Exit Sub
    
    Application.ScreenUpdating = False
    
    For colIdx = colStart To colEnd
        For rowIdx = rowStart + 1 To lastRow
            currVal = ws.Cells(rowIdx, colIdx).Value
            prevVal = ws.Cells(rowIdx - 1, colIdx).Value
            
            If Not IsEmpty(currVal) And Not IsEmpty(prevVal) Then
                If currVal = prevVal Then
                    ws.Cells(rowIdx, colIdx).Interior.Color = highlightColor
                    ws.Cells(rowIdx - 1, colIdx).Interior.Color = highlightColor
                End If
            End If
        Next rowIdx
    Next colIdx
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_HighlightDuplicates 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_HighlightDupSingleCol
' 单列重复高亮便捷方法
' 参数:
'   ws             — 目标工作表
'   col            — 列号（1-based），默认 5（E列）
'   rowStart       — 数据起始行号，默认 2
'   highlightColor — 高亮颜色，默认 vbYellow
'============================================================================
Public Sub DM_HighlightDupSingleCol(ByVal ws As Worksheet, _
    Optional ByVal col As Long = 5, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal highlightColor As Long = 65535)
    
    Call DM_HighlightDuplicates(ws, col, col, rowStart, highlightColor)
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 3: modBlankRows.bas — 空白行管理

**Files:**
- Create: `vba/modules/modBlankRows.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modBlankRows"
Option Explicit

'============================================================================
' 模块: modBlankRows
' 描述: CDM空白行管理 — 按规则插入/删除空白行
'        Blank Row Management — insert blank rows every N rows, between groups, or delete empties
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

'============================================================================
' DM_InsertBlankEveryN
' 每N行插入一行空白（自底向上执行，保持行号正确）
' 参数:
'   ws        — 目标工作表
'   n         — 间隔行数，默认 5
'   colLetter — 用于确定数据最后行的列字母，默认 "A"
' 返回: Long — 插入的空白行数
' 用法:
'   Dim n As Long: n = DM_InsertBlankEveryN(ActiveSheet, 5, "A")
'============================================================================
Public Function DM_InsertBlankEveryN(ByVal ws As Worksheet, _
    Optional ByVal n As Long = 5, _
    Optional ByVal colLetter As String = "A") As Long
    
    On Error GoTo ErrHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim inserted As Long
    
    If n < 1 Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.count, colLetter).End(xlUp).Row
    If lastRow <= 1 Then Exit Function
    
    Application.ScreenUpdating = False
    
    For i = lastRow To 2 Step -1
        If (i - 1) Mod n = 0 Then
            ws.Rows(i + 1).Insert Shift:=xlDown
            inserted = inserted + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    DM_InsertBlankEveryN = inserted
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    DM_InsertBlankEveryN = -1
    MsgBox "DM_InsertBlankEveryN 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_InsertBlankBetweenGroups
' 当指定列的值变化时在组间插入空白行
' 参数:
'   ws        — 目标工作表
'   colLetter — 用于判断分组的列字母，默认 "A"
' 返回: Long — 插入的空白行数
'============================================================================
Public Function DM_InsertBlankBetweenGroups(ByVal ws As Worksheet, _
    Optional ByVal colLetter As String = "A") As Long
    
    On Error GoTo ErrHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim inserted As Long
    Dim currVal As Variant
    Dim nextVal As Variant
    
    lastRow = ws.Cells(ws.Rows.count, colLetter).End(xlUp).Row
    If lastRow <= 2 Then Exit Function
    
    Application.ScreenUpdating = False
    
    For i = lastRow To 2 Step -1
        currVal = ws.Cells(i, colLetter).Value
        nextVal = ws.Cells(i + 1, colLetter).Value
        
        If Not IsEmpty(currVal) And Not IsEmpty(nextVal) Then
            If currVal <> nextVal Then
                ws.Rows(i + 1).Insert Shift:=xlDown
                inserted = inserted + 1
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    DM_InsertBlankBetweenGroups = inserted
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    DM_InsertBlankBetweenGroups = -1
    MsgBox "DM_InsertBlankBetweenGroups 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_DeleteBlankRows
' 删除工作表中完全空白的行（所有单元格均为空）
' 参数:
'   ws — 目标工作表
' 返回: Long — 删除的行数
'============================================================================
Public Function DM_DeleteBlankRows(ByVal ws As Worksheet) As Long
    
    On Error GoTo ErrHandler
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim isEmptyRow As Boolean
    Dim deleted As Long
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    If lastRow <= 1 Then Exit Function
    If lastCol < 1 Then lastCol = 1
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = lastRow To 1 Step -1
        isEmptyRow = True
        For j = 1 To lastCol
            If Not IsEmpty(ws.Cells(i, j).Value) Then
                isEmptyRow = False
                Exit For
            End If
        Next j
        If isEmptyRow Then
            ws.Rows(i).Delete Shift:=xlUp
            deleted = deleted + 1
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_DeleteBlankRows = deleted
    Exit Function
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_DeleteBlankRows = -1
    MsgBox "DM_DeleteBlankRows 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function
```

- [ ] **Step 2: 验证编译**

---

## 阶段二：核心模块（有原始VBA参考，中等复杂度）

### Task 4: modTOC.bas — 目录生成器

**Files:**
- Create: `vba/modules/modTOC.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modTOC"
Option Explicit

'============================================================================
' 模块: modTOC
' 描述: CDM目录生成器 — 生成带超链接的工作簿目录（标准版 + SAS Derive版）
'        Table of Contents generator with hyperlinks + back-arrow navigation
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

Private Const BACK_ARROW_NAME As String = "CDM_BackToTOC"

'============================================================================
' DM_SheetExists
' 判断指定名称的工作表是否存在
'============================================================================
Private Function DM_SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    DM_SheetExists = (Err.Number = 0)
    Err.Clear
End Function

'============================================================================
' DM_GenerateTOC
' 创建/重建目录sheet，包含指向所有可见sheet的超链接和返回箭头
' 参数:
'   wb            — 目标工作簿
'   tocName       — 目录sheet名称，默认 "TOC"
'   excludeSheets — 逗号分隔的排除sheet列表，如 "TOC,Settings"
' 用法:
'   DM_GenerateTOC ActiveWorkbook
'   DM_GenerateTOC ActiveWorkbook, "目录", "Settings,Temp"
'============================================================================
Public Sub DM_GenerateTOC(ByVal wb As Workbook, _
    Optional ByVal tocName As String = "TOC", _
    Optional ByVal excludeSheets As String = "")
    
    On Error GoTo ErrHandler
    
    Dim toc As Worksheet
    Dim ws As Worksheet
    Dim wsName As Variant
    Dim excludeArr() As String
    Dim rowIdx As Long
    Dim cell As Range
    Dim shp As Shape
    Dim isExcluded As Boolean
    Dim i As Long
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 解析排除列表
    If Len(excludeSheets) > 0 Then
        excludeArr = Split(excludeSheets, ",")
    End If
    
    ' 删除已有目录sheet（如果存在）
    If DM_SheetExists(wb, tocName) Then
        wb.Worksheets(tocName).Delete
    End If
    
    ' 创建新目录sheet（放在最前面）
    Set toc = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    toc.Name = tocName
    toc.Range("A1").Value = "Table of Contents"
    toc.Range("A1").Font.Bold = True
    toc.Range("A1").Font.Size = 14
    toc.Range("A1").Interior.Color = RGB(68, 114, 196)  ' 蓝色背景
    toc.Range("A1").Font.Color = vbWhite
    
    rowIdx = 2
    
    For Each ws In wb.Worksheets
        If ws.Name <> tocName Then
            ' 检查是否在排除列表中
            isExcluded = False
            If Len(excludeSheets) > 0 Then
                For i = 0 To UBound(excludeArr)
                    If Trim(ws.Name) = Trim(excludeArr(i)) Then
                        isExcluded = True
                        Exit For
                    End If
                Next i
            End If
            
            If Not isExcluded Then
                ' 列A: 带超链接的sheet名称
                Set cell = toc.Cells(rowIdx, 1)
                cell.Value = ws.Name
                cell.Font.Color = RGB(5, 99, 193)  ' 超链接蓝
                cell.Font.Underline = xlUnderlineStyleSingle
                toc.Hyperlinks.Add Anchor:=cell, Address:="", _
                    SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
                
                ' 删除旧返回箭头（如果存在）
                On Error Resume Next
                ws.Shapes(BACK_ARROW_NAME).Delete
                On Error GoTo ErrHandler
                
                ' 添加返回箭头
                Set shp = ws.Shapes.AddShape(msoShapeLeftArrow, 0, 0, 20, 10)
                shp.Name = BACK_ARROW_NAME
                shp.TextFrame2.TextRange.Text = "←"
                ws.Hyperlinks.Add Anchor:=shp, Address:="", _
                    SubAddress:="'" & tocName & "'!A1"
                
                rowIdx = rowIdx + 1
            End If
        End If
    Next ws
    
    toc.Columns("A").ColumnWidth = 35
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "DM_GenerateTOC 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_GenerateDeriveTOC
' SAS Derive风格目录：A列=超链接sheet名称，B列=从指定单元格提取的描述
' 参数:
'   wb       — 目标工作簿
'   descCell — 每个sheet中提取描述的单元格地址，默认 "V2"
'   tocName  — 目录sheet名称，默认 "TOC"
' 用法:
'   DM_GenerateDeriveTOC ActiveWorkbook, "V2", "TOC"
'============================================================================
Public Sub DM_GenerateDeriveTOC(ByVal wb As Workbook, _
    Optional ByVal descCell As String = "V2", _
    Optional ByVal tocName As String = "TOC")
    
    On Error GoTo ErrHandler
    
    Dim toc As Worksheet
    Dim ws As Worksheet
    Dim rowIdx As Long
    Dim cellA As Range
    Dim cellB As Range
    Dim desc As Variant
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 删除已有目录sheet
    If DM_SheetExists(wb, tocName) Then
        wb.Worksheets(tocName).Delete
    End If
    
    ' 创建新目录sheet
    Set toc = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    toc.Name = tocName
    
    ' 标题行
    toc.Range("A1").Value = "TOC"
    toc.Range("A1").Font.Bold = True
    toc.Range("A1").Font.Size = 14
    toc.Range("B1").Value = "Description"
    toc.Range("B1").Font.Bold = True
    toc.Range("B1").Font.Size = 14
    
    rowIdx = 2
    
    For Each ws In wb.Worksheets
        If ws.Name <> tocName And ws.Visible = xlSheetVisible Then
            ' A列: 超链接sheet名称
            Set cellA = toc.Cells(rowIdx, 1)
            cellA.Value = ws.Name
            cellA.Font.Color = RGB(5, 99, 193)
            cellA.Font.Underline = xlUnderlineStyleSingle
            toc.Hyperlinks.Add Anchor:=cellA, Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            
            ' B列: 从指定单元格提取描述
            On Error Resume Next
            desc = ws.Range(descCell).Value
            On Error GoTo ErrHandler
            Set cellB = toc.Cells(rowIdx, 2)
            If Not IsEmpty(desc) Then
                cellB.Value = CStr(desc)
            End If
            
            rowIdx = rowIdx + 1
        End If
    Next ws
    
    toc.Columns("A").ColumnWidth = 35
    toc.Columns("B").ColumnWidth = 50
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "DM_GenerateDeriveTOC 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 5: modSheetSplit.bas — 工作表拆分

**Files:**
- Create: `vba/modules/modSheetSplit.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modSheetSplit"
Option Explicit

'============================================================================
' 模块: modSheetSplit
' 描述: CDM工作表拆分 — 按列值将一个sheet拆分为多个sheet
'        Sheet Splitter — split one sheet into many by column value
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

'============================================================================
' DM_SplitSheetByColumn
' 按指定列的唯一值将工作表拆分为多个sheet
' 参数:
'   ws         — 源工作表
'   splitCol   — 拆分键所在的列字母，默认 "A"
'   headerRows — 表头行数，将复制到每个子sheet顶部，默认 1
' 返回: String() — 创建的sheet名称数组
' 用法:
'   Dim sheets() As String
'   sheets = DM_SplitSheetByColumn(ActiveSheet, "A", 1)
'============================================================================
Public Function DM_SplitSheetByColumn(ByVal ws As Worksheet, _
    Optional ByVal splitCol As String = "A", _
    Optional ByVal headerRows As Long = 1) As String()
    
    On Error GoTo ErrHandler
    
    Dim wb As Workbook
    Dim lastRow As Long
    Dim lastCol As Long
    Dim splitColIdx As Long
    Dim i As Long
    Dim key As Variant
    Dim groups As Object ' Scripting.Dictionary: key → Collection of row numbers
    Dim rowColl As Collection
    Dim child As Worksheet
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim srcRow As Long
    Dim destRow As Long
    Dim result() As String
    Dim k As Long
    Dim wsIdx As Long
    
    Set wb = ws.Parent
    splitColIdx = ws.Range(splitCol & "1").Column
    lastRow = ws.Cells(ws.Rows.count, splitColIdx).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    If lastRow <= headerRows Then
        ReDim result(0 To -1)
        DM_SplitSheetByColumn = result
        Exit Function
    End If
    
    ' 按键值分组（记录行号）
    Set groups = CreateObject("Scripting.Dictionary")
    For i = headerRows + 1 To lastRow
        key = CStr(ws.Cells(i, splitColIdx).Value)
        If Len(key) = 0 Then key = "(blank)"
        If Not groups.Exists(key) Then
            Set rowColl = New Collection
            groups.Add rowColl, key
        End If
        groups(key).Add i
    Next i
    
    ' 检测sheet名称冲突
    wsIdx = ws.Index
    Dim conflictNames As String
    For Each key In groups.Keys
        Dim safeName As String
        safeName = DM_MakeSafeSheetName(CStr(key))
        If DM_SheetExistsLocal(wb, safeName) Then
            conflictNames = conflictNames & safeName & vbCrLf
        End If
    Next key
    If Len(conflictNames) > 0 Then
        MsgBox "以下sheet名称已存在，请先删除或重命名:" & vbCrLf & conflictNames, _
            vbExclamation, "CDM Toolkit - 拆分冲突"
        ReDim result(0 To -1)
        DM_SplitSheetByColumn = result
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ReDim result(0 To groups.count - 1)
    k = 0
    
    For Each key In groups.Keys
        safeName = DM_MakeSafeSheetName(CStr(key))
        Set child = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        child.Name = safeName
        result(k) = safeName
        k = k + 1
        
        destRow = 1
        
        ' 复制表头行
        If headerRows > 0 Then
            For srcRow = 1 To headerRows
                For colIdx = 1 To lastCol
                    child.Cells(destRow, colIdx).Value = ws.Cells(srcRow, colIdx).Value
                    child.Cells(destRow, colIdx).Font.Bold = True
                Next colIdx
                destRow = destRow + 1
            Next srcRow
        End If
        
        ' 复制数据行
        Set rowColl = groups(key)
        Dim v As Variant
        For Each v In rowColl
            srcRow = CLng(v)
            For colIdx = 1 To lastCol
                child.Cells(destRow, colIdx).Value = ws.Cells(srcRow, colIdx).Value
            Next colIdx
            destRow = destRow + 1
        Next v
    Next key
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_SplitSheetByColumn = result
    Exit Function
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "DM_SplitSheetByColumn 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_SplitSheetWithProgress
' 带进度指示器的拆分方法，适用于大数据量
'============================================================================
Public Function DM_SplitSheetWithProgress(ByVal ws As Worksheet, _
    Optional ByVal splitCol As String = "A", _
    Optional ByVal headerRows As Long = 1) As String()
    
    Dim oldStatusBar As String
    oldStatusBar = Application.StatusBar
    Application.StatusBar = "CDM Toolkit: 正在拆分工作表..."
    DM_SplitSheetWithProgress = DM_SplitSheetByColumn(ws, splitCol, headerRows)
    Application.StatusBar = oldStatusBar
End Function

'============================================================================
' 辅助函数
'============================================================================

' 生成安全的sheet名称（去除非法字符，截断到31字符）
Private Function DM_MakeSafeSheetName(ByVal name As String) As String
    Dim result As String
    Dim illegalChars As String
    Dim i As Long
    
    illegalChars = "\/*?:[]"
    result = name
    
    For i = 1 To Len(illegalChars)
        result = Replace(result, Mid(illegalChars, i, 1), "_")
    Next i
    
    If Len(result) > 31 Then result = Left(result, 31)
    If Len(result) = 0 Then result = "Sheet"
    DM_MakeSafeSheetName = result
End Function

' 检查sheet是否存在（模块内私有）
Private Function DM_SheetExistsLocal(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    DM_SheetExistsLocal = (Err.Number = 0)
    Err.Clear
End Function
```

- [ ] **Step 2: 验证编译**

---

### Task 6: modWorkbookMerge.bas — 工作簿合并

**Files:**
- Create: `vba/modules/modWorkbookMerge.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modWorkbookMerge"
Option Explicit

'============================================================================
' 模块: modWorkbookMerge
' 描述: CDM工作簿合并 — 将多个Excel工作簿合并为一个（含可选TOC）
'        Workbook Merger — merge multiple workbooks into one with optional TOC
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: modTOC.bas（需先导入）
'============================================================================

'============================================================================
' DM_MergeWorkbooksDialog
' 通过文件对话框选择文件后合并
' 参数:
'   createTOC — 是否在合并后生成TOC目录，默认 True
'============================================================================
Public Sub DM_MergeWorkbooksDialog(Optional ByVal createTOC As Boolean = True)
    On Error GoTo ErrHandler
    
    Dim fd As FileDialog
    Dim fileList As Collection
    Dim v As Variant
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择要合并的Excel文件"
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xlsx;*.xls", 1
        .AllowMultiSelect = True
        
        If .Show <> -1 Then Exit Sub
        
        Set fileList = New Collection
        For Each v In .SelectedItems
            fileList.Add CStr(v)
        Next v
    End With
    
    DM_MergeWorkbooksCore fileList, createTOC
    Exit Sub
    
ErrHandler:
    MsgBox "DM_MergeWorkbooksDialog 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_MergeWorkbooksFolder
' 合并指定文件夹中所有匹配的文件
' 参数:
'   folderPath  — 文件夹路径
'   filePattern — 文件匹配模式，默认 "*.xlsx"
'   createTOC   — 是否生成TOC，默认 True
'============================================================================
Public Sub DM_MergeWorkbooksFolder(ByVal folderPath As String, _
    Optional ByVal filePattern As String = "*.xlsx", _
    Optional ByVal createTOC As Boolean = True)
    
    On Error GoTo ErrHandler
    
    Dim fileName As String
    Dim fileList As Collection
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Set fileList = New Collection
    fileName = Dir(folderPath & filePattern)
    Do While fileName <> ""
        fileList.Add folderPath & fileName
        fileName = Dir()
    Loop
    
    If fileList.count = 0 Then
        MsgBox "在 " & folderPath & " 中未找到匹配 " & filePattern & " 的文件。", _
            vbInformation, "CDM Toolkit"
        Exit Sub
    End If
    
    DM_MergeWorkbooksCore fileList, createTOC
    Exit Sub
    
ErrHandler:
    MsgBox "DM_MergeWorkbooksFolder 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_MergeWorkbooksList
' 合并指定文件路径列表中的工作簿
'============================================================================
Public Sub DM_MergeWorkbooksList(ByVal filePaths As Collection, _
    Optional ByVal createTOC As Boolean = True)
    DM_MergeWorkbooksCore filePaths, createTOC
End Sub

'============================================================================
' DM_MergeWorkbooksCore
' 核心合并逻辑
'============================================================================
Private Sub DM_MergeWorkbooksCore(ByVal fileList As Collection, _
    ByVal createTOC As Boolean)
    
    On Error GoTo ErrHandler
    
    Dim newWb As Workbook
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim newWs As Worksheet
    Dim filePath As Variant
    Dim fileName As String
    Dim prefix As String
    Dim newName As String
    Dim counter As Long
    Dim i As Long
    Dim baseName As String
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 创建目标工作簿
    Set newWb = Workbooks.Add
    
    For Each filePath In fileList
        Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True)
        fileName = CreateObject("Scripting.FileSystemObject").GetBaseName(CStr(filePath))
        prefix = Left(fileName, 20)  ' 文件名前缀（截断）
        
        For Each srcWs In srcWb.Worksheets
            newName = prefix & "_" & srcWs.Name
            If Len(newName) > 31 Then newName = Left(newName, 31)
            
            ' 处理重名冲突
            counter = 1
            baseName = newName
            Do While DM_SheetExistsInWorkbook(newWb, newName)
                newName = Left(baseName, 27) & "_" & counter
                counter = counter + 1
            Loop
            
            ' 复制工作表
            srcWs.Copy Before:=newWb.Worksheets(newWb.Worksheets.count + 1)
            Set newWs = newWb.Worksheets(newWb.Worksheets.count)
            newWs.Name = newName
        Next srcWs
        
        srcWb.Close SaveChanges:=False
    Next filePath
    
    ' 删除默认空sheet
    On Error Resume Next
    newWb.Worksheets("Sheet1").Delete
    On Error GoTo ErrHandler
    
    ' 生成TOC
    If createTOC Then
        Call DM_MergeGenerateTOC(newWb)
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "合并完成！共合并了 " & fileList.count & " 个工作簿，" & _
        newWb.Worksheets.count & " 个sheet。", vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "DM_MergeWorkbooksCore 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' 辅助函数
'============================================================================

Private Function DM_SheetExistsInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    DM_SheetExistsInWorkbook = (Err.Number = 0)
    Err.Clear
End Function

' 内联TOC生成（避免硬依赖 modTOC）
Private Sub DM_MergeGenerateTOC(ByVal wb As Workbook)
    On Error Resume Next
    Dim toc As Worksheet
    Dim ws As Worksheet
    Dim rowIdx As Long
    Dim cell As Range
    
    If DM_SheetExistsInWorkbook(wb, "TOC") Then
        wb.Worksheets("TOC").Delete
    End If
    
    Set toc = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    toc.Name = "TOC"
    toc.Range("A1").Value = "Table of Contents"
    toc.Range("A1").Font.Bold = True
    toc.Range("A1").Font.Size = 14
    
    rowIdx = 2
    For Each ws In wb.Worksheets
        If ws.Name <> "TOC" Then
            Set cell = toc.Cells(rowIdx, 1)
            cell.Value = ws.Name
            toc.Hyperlinks.Add Anchor:=cell, Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            rowIdx = rowIdx + 1
        End If
    Next ws
    toc.Columns("A").ColumnWidth = 35
    On Error GoTo 0
End Sub
```

- [ ] **Step 2: 验证编译**

---

## 阶段三：全新编写（从Python反向编写，无原始VBA参考）

### Task 7: modSheetManager.bas — 工作表管理

**Files:**
- Create: `vba/modules/modSheetManager.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modSheetManager"
Option Explicit

'============================================================================
' 模块: modSheetManager
' 描述: CDM工作表管理 — 批量重命名、名称列表、删除隐藏/指定sheet
'        Sheet Manager — batch rename, list names, delete hidden/specified sheets
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

'============================================================================
' DM_ListSheetNames
' 获取工作簿中所有工作表名称
' 参数:
'   wb            — 目标工作簿
'   includeHidden — 是否包含隐藏sheet，默认 False
' 返回: String() — sheet名称数组
' 用法:
'   Dim names() As String
'   names = DM_ListSheetNames(ActiveWorkbook)
'   Dim n As Long: For n = 0 To UBound(names): Debug.Print names(n): Next n
'============================================================================
Public Function DM_ListSheetNames(ByVal wb As Workbook, _
    Optional ByVal includeHidden As Boolean = False) As String()
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim count As Long
    Dim result() As String
    Dim i As Long
    
    ' 先统计数量
    count = 0
    For Each ws In wb.Worksheets
        If includeHidden Or ws.Visible = xlSheetVisible Then
            count = count + 1
        End If
    Next ws
    
    If count = 0 Then
        ReDim result(0 To -1)
        DM_ListSheetNames = result
        Exit Function
    End If
    
    ReDim result(0 To count - 1)
    i = 0
    For Each ws In wb.Worksheets
        If includeHidden Or ws.Visible = xlSheetVisible Then
            result(i) = ws.Name
            i = i + 1
        End If
    Next ws
    
    DM_ListSheetNames = result
    Exit Function
    
ErrHandler:
    MsgBox "DM_ListSheetNames 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_BatchRenameSheets
' 批量重命名工作表
' 参数:
'   wb      — 目标工作簿
'   nameMap — Dictionary {旧名称: 新名称}
' 返回: Long — 成功重命名的sheet数量
' 用法:
'   Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
'   map.Add "AE", "Adverse_Events"
'   map.Add "LB", "Laboratory"
'   DM_BatchRenameSheets ActiveWorkbook, map
'============================================================================
Public Function DM_BatchRenameSheets(ByVal wb As Workbook, _
    ByVal nameMap As Object) As Long
    
    On Error GoTo ErrHandler
    
    Dim key As Variant
    Dim newName As String
    Dim count As Long
    
    If nameMap Is Nothing Then Exit Function
    If nameMap.count = 0 Then Exit Function
    
    Application.ScreenUpdating = False
    
    For Each key In nameMap.Keys
        newName = CStr(nameMap(key))
        If Len(newName) > 31 Then newName = Left(newName, 31)
        
        On Error Resume Next
        wb.Worksheets(CStr(key)).Name = newName
        If Err.Number = 0 Then count = count + 1
        Err.Clear
        On Error GoTo ErrHandler
    Next key
    
    Application.ScreenUpdating = True
    DM_BatchRenameSheets = count
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_BatchRenameSheets 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_DeleteHiddenSheets
' 删除所有隐藏工作表
' 参数:
'   wb — 目标工作簿
' 返回: String() — 被删除的sheet名称数组
'============================================================================
Public Function DM_DeleteHiddenSheets(ByVal wb As Workbook) As String()
    
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim count As Long
    Dim result() As String
    Dim i As Long
    
    ' 统计隐藏sheet数量
    count = 0
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            count = count + 1
        End If
    Next ws
    
    If count = 0 Then
        ReDim result(0 To -1)
        DM_DeleteHiddenSheets = result
        Exit Function
    End If
    
    ReDim result(0 To count - 1)
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 需要反向遍历，因为删除会改变集合
    i = 0
    Dim j As Long
    For j = wb.Worksheets.count To 1 Step -1
        Set ws = wb.Worksheets(j)
        If ws.Visible <> xlSheetVisible Then
            result(i) = ws.Name
            ws.Delete
            i = i + 1
        End If
    Next j
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    DM_DeleteHiddenSheets = result
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "DM_DeleteHiddenSheets 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_DeleteSheets
' 删除指定名称的工作表（带确认）
' 参数:
'   wb         — 目标工作簿
'   sheetNames — 可变参数，要删除的sheet名称列表
' 用法:
'   DM_DeleteSheets ActiveWorkbook, "Temp1", "Temp2"
'============================================================================
Public Sub DM_DeleteSheets(ByVal wb As Workbook, ParamArray sheetNames() As Variant)
    
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim sheetName As String
    Dim sheetList As String
    Dim ws As Worksheet
    
    If UBound(sheetNames) < 0 Then Exit Sub
    
    ' 构建sheet名称列表用于确认
    For i = 0 To UBound(sheetNames)
        sheetList = sheetList & vbCrLf & "  - " & CStr(sheetNames(i))
    Next i
    
    If MsgBox("确认删除以下工作表？" & sheetList, vbYesNo + vbQuestion, "CDM Toolkit - 确认删除") <> vbYes Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    For i = 0 To UBound(sheetNames)
        sheetName = CStr(sheetNames(i))
        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        If Err.Number = 0 Then
            ws.Delete
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "DM_DeleteSheets 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 8: modDateCalc.bas — 日期计算器

**Files:**
- Create: `vba/modules/modDateCalc.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modDateCalc"
Option Explicit

'============================================================================
' 模块: modDateCalc
' 描述: CDM日期计算器 — 日期列加减天数、计算两列日期差
'        Date Calculator — add/subtract days, compute date differences
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

'============================================================================
' DM_AddDaysToColumn
' 对指定列中每个日期值加/减天数
' 参数:
'   ws         — 目标工作表
'   colLetter  — 目标列字母
'   days       — 要加的天数（负数表示减）
'   rowStart   — 数据起始行，默认 2
'   dateFormat — 结果日期格式，默认 "YYYY-MM-DD"
' 返回: Long — 更新的单元格数量
' 用法:
'   DM_AddDaysToColumn ActiveSheet, "F", 30          ' F列所有日期 +30天
'   DM_AddDaysToColumn ActiveSheet, "G", -7, 2, "DD/MM/YYYY"
'============================================================================
Public Function DM_AddDaysToColumn(ByVal ws As Worksheet, _
    ByVal colLetter As String, _
    ByVal days As Long, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal dateFormat As String = "YYYY-MM-DD") As Long
    
    On Error GoTo ErrHandler
    
    Dim col As Long
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim count As Long
    
    col = ws.Range(colLetter & "1").Column
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    
    If lastRow < rowStart Then Exit Function
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = rowStart To lastRow
        Set cell = ws.Cells(i, col)
        If IsDate(cell.Value) Then
            cell.Value = DateAdd("d", days, CDate(cell.Value))
            cell.NumberFormat = dateFormat
            count = count + 1
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_AddDaysToColumn = count
    Exit Function
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_AddDaysToColumn = -1
    MsgBox "DM_AddDaysToColumn 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_CalcDateDiff
' 计算两列日期之间的天数差，写入结果列
' 参数:
'   ws        — 目标工作表
'   colStart  — 起始日期列字母
'   colEnd    — 结束日期列字母
'   resultCol — 结果列字母
'   rowStart  — 数据起始行，默认 2
' 返回: Long — 计算的单元格数量
' 用法:
'   DM_CalcDateDiff ActiveSheet, "F", "G", "M"  ' (G列-F列)天数 → M列
'============================================================================
Public Function DM_CalcDateDiff(ByVal ws As Worksheet, _
    ByVal colStart As String, _
    ByVal colEnd As String, _
    ByVal resultCol As String, _
    Optional ByVal rowStart As Long = 2) As Long
    
    On Error GoTo ErrHandler
    
    Dim c1 As Long, c2 As Long, cr As Long
    Dim lastRow As Long
    Dim i As Long
    Dim d1 As Date, d2 As Date
    Dim count As Long
    
    c1 = ws.Range(colStart & "1").Column
    c2 = ws.Range(colEnd & "1").Column
    cr = ws.Range(resultCol & "1").Column
    lastRow = ws.Cells(ws.Rows.count, c1).End(xlUp).Row
    
    If lastRow < rowStart Then Exit Function
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = rowStart To lastRow
        If IsDate(ws.Cells(i, c1).Value) And IsDate(ws.Cells(i, c2).Value) Then
            d1 = CDate(ws.Cells(i, c1).Value)
            d2 = CDate(ws.Cells(i, c2).Value)
            ws.Cells(i, cr).Value = d2 - d1
            count = count + 1
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_CalcDateDiff = count
    Exit Function
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_CalcDateDiff = -1
    MsgBox "DM_CalcDateDiff 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_DateAddSelection
' 便捷方法：对当前选中的所有单元格加天数
' 参数:
'   days — 要加的天数（负数表示减）
'============================================================================
Public Sub DM_DateAddSelection(ByVal days As Long)
    
    On Error GoTo ErrHandler
    
    Dim cell As Range
    Dim count As Long
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Application.ScreenUpdating = False
    
    For Each cell In Selection
        If IsDate(cell.Value) Then
            cell.Value = DateAdd("d", days, CDate(cell.Value))
            count = count + 1
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "已更新 " & count & " 个日期单元格。", vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_DateAddSelection 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 9: modUnitConvert.bas — 实验室单位转换

**Files:**
- Create: `vba/modules/modUnitConvert.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modUnitConvert"
Option Explicit

'============================================================================
' 模块: modUnitConvert
' 描述: CDM实验室单位转换 — 将实验室结果统一转换为标准单位
'        Lab Unit Conversion — convert lab results to a standard unit
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

' --- 模块级转换表 ---
Private mConversionTable As Object ' Scripting.Dictionary: "fromUnit→toUnit" → factor

'============================================================================
' DM_InitConversionTable
' 初始化内置转换因子表（与 Python UNIT_CONVERSIONS 完全对齐）
'============================================================================
Private Sub DM_InitConversionTable()
    If Not mConversionTable Is Nothing Then Exit Sub
    
    Set mConversionTable = CreateObject("Scripting.Dictionary")
    
    ' 质量 Mass
    DM_AddConv "g", "mg", 1000
    DM_AddConv "g", "ug", 1000000
    DM_AddConv "mg", "g", 0.001
    DM_AddConv "mg", "ug", 1000
    DM_AddConv "ug", "g", 0.000001
    DM_AddConv "ug", "mg", 0.001
    DM_AddConv "kg", "g", 1000
    DM_AddConv "g", "kg", 0.001
    
    ' 体积 Volume
    DM_AddConv "L", "mL", 1000
    DM_AddConv "mL", "L", 0.001
    DM_AddConv "dL", "L", 0.1
    DM_AddConv "L", "dL", 10
    
    ' 长度 Length
    DM_AddConv "cm", "mm", 10
    DM_AddConv "mm", "cm", 0.1
    DM_AddConv "m", "cm", 100
    DM_AddConv "cm", "m", 0.01
    
    ' 浓度 Concentration
    DM_AddConv "g/L", "mg/mL", 1
    DM_AddConv "mg/mL", "g/L", 1
    DM_AddConv "mg/dL", "g/L", 0.01
    DM_AddConv "g/L", "mg/dL", 100
    DM_AddConv "g/dL", "g/L", 10
    DM_AddConv "g/L", "g/dL", 0.1
    DM_AddConv "umol/L", "mg/dL", 0.0113   ' 肌酐近似值
    DM_AddConv "mg/dL", "umol/L", 88.42
    
    ' 时间 Time
    DM_AddConv "h", "min", 60
    DM_AddConv "min", "h", 1 / 60
    DM_AddConv "d", "h", 24
    DM_AddConv "h", "d", 1 / 24
End Sub

' 内部：添加单个转换
Private Sub DM_AddConv(ByVal fromUnit As String, ByVal toUnit As String, ByVal factor As Double)
    Dim key As String
    key = LCase(fromUnit) & "→" & LCase(toUnit)
    If Not mConversionTable.Exists(key) Then
        mConversionTable.Add key, factor
    End If
End Sub

'============================================================================
' DM_ConvertLabUnits
' 将实验室结果转换为标准单位
' 参数:
'   ws            — 目标工作表
'   testCol       — 检验项目列字母（如 "D"）
'   resultCol     — 结果值列字母（如 "E"）
'   unitCol       — 单位列字母（如 "F"）
'   standardUnit  — 目标标准单位（如 "mg/dL"）
'   testFilter    — 可选，仅转换此检验项目（如 "Glucose"）
' 返回: Long — 成功转换的单元格数量
' 用法:
'   DM_ConvertLabUnits ActiveSheet, "D", "E", "F", "mg/dL", "Glucose"
'============================================================================
Public Function DM_ConvertLabUnits(ByVal ws As Worksheet, _
    ByVal testCol As String, _
    ByVal resultCol As String, _
    ByVal unitCol As String, _
    ByVal standardUnit As String, _
    Optional ByVal testFilter As String = "") As Long
    
    On Error GoTo ErrHandler
    
    Call DM_InitConversionTable
    
    Dim tc As Long, rc As Long, uc As Long
    Dim lastRow As Long
    Dim i As Long
    Dim testName As String
    Dim unitVal As String
    Dim resultVal As Variant
    Dim factor As Variant
    Dim key As String
    Dim count As Long
    
    tc = ws.Range(testCol & "1").Column
    rc = ws.Range(resultCol & "1").Column
    uc = ws.Range(unitCol & "1").Column
    lastRow = ws.Cells(ws.Rows.count, tc).End(xlUp).Row
    
    If lastRow < 2 Then Exit Function
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 2 To lastRow
        testName = CStr(ws.Cells(i, tc).Value)
        unitVal = CStr(ws.Cells(i, uc).Value)
        resultVal = ws.Cells(i, rc).Value
        
        ' 检查检验项目过滤
        If Len(testFilter) > 0 And LCase(testName) <> LCase(testFilter) Then GoTo NextRow
        
        ' 单位已经是目标单位，跳过
        If LCase(unitVal) = LCase(standardUnit) Then GoTo NextRow
        
        ' 检查结果是否为数值
        If Not IsNumeric(resultVal) Then GoTo NextRow
        
        ' 查找转换因子
        key = LCase(unitVal) & "→" & LCase(standardUnit)
        If mConversionTable.Exists(key) Then
            factor = mConversionTable(key)
            ws.Cells(i, rc).Value = Round(CDbl(resultVal) * CDbl(factor), 6)
            ws.Cells(i, uc).Value = standardUnit
            count = count + 1
        Else
            ' 标记无法转换的行为红色
            ws.Cells(i, rc).Interior.Color = RGB(255, 153, 153)
        End If
        
NextRow:
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_ConvertLabUnits = count
    Exit Function
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_ConvertLabUnits = -1
    MsgBox "DM_ConvertLabUnits 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_GetSupportedConversions
' 列出所有可用的单位转换
'============================================================================
Public Function DM_GetSupportedConversions() As String()
    Call DM_InitConversionTable
    
    Dim result() As String
    Dim key As Variant
    Dim i As Long
    
    ReDim result(0 To mConversionTable.count - 1)
    i = 0
    For Each key In mConversionTable.Keys
        result(i) = CStr(key) & " = " & mConversionTable(key)
        i = i + 1
    Next key
    DM_GetSupportedConversions = result
End Function

'============================================================================
' DM_AddUnitConversion
' 在运行时扩展转换表
' 参数:
'   fromUnit — 源单位
'   toUnit   — 目标单位
'   factor   — 转换因子 (结果 = 原始值 × factor)
'============================================================================
Public Sub DM_AddUnitConversion(ByVal fromUnit As String, _
    ByVal toUnit As String, ByVal factor As Double)
    
    Call DM_InitConversionTable
    DM_AddConv fromUnit, toUnit, factor
End Sub
```

- [ ] **Step 2: 验证编译**

---

## 阶段四：高级模块（更复杂的逻辑）

### Task 10: modTrackChanges.bas — 修改痕迹跟踪

**Files:**
- Create: `vba/modules/modTrackChanges.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modTrackChanges"
Option Explicit

'============================================================================
' 模块: modTrackChanges
' 描述: CDM修改痕迹跟踪 — 比较工作表与原版，通过批注+颜色标记修改
'        Track Changes — compare worksheet against original, mark changes
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

'============================================================================
' DM_TrackChanges
' 比较当前工作表与原始备份副本，标记所有修改过的单元格
' 参数:
'   ws               — 当前工作表
'   originalFilePath — 原始工作簿文件路径
'   addComments      — 是否添加批注显示"旧值→新值"，默认 True
'   highlightChanges — 是否高亮修改过的单元格（绿填充+红边框），默认 True
' 返回: Long — 检测到的修改数量
' 用法:
'   DM_TrackChanges ActiveSheet, "C:\backup\original.xlsx"
'============================================================================
Public Function DM_TrackChanges(ByVal ws As Worksheet, _
    ByVal originalFilePath As String, _
    Optional ByVal addComments As Boolean = True, _
    Optional ByVal highlightChanges As Boolean = True) As Long
    
    On Error GoTo ErrHandler
    
    Dim origWb As Workbook
    Dim origWs As Worksheet
    Dim cell As Range
    Dim origVal As Variant
    Dim changes As Long
    Dim cmt As Comment
    Dim sheetName As String
    
    sheetName = ws.Name
    
    ' 打开原始工作簿
    If Dir(originalFilePath) = "" Then
        MsgBox "原始文件未找到: " & originalFilePath, vbExclamation, "CDM Toolkit"
        DM_TrackChanges = -1
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set origWb = Workbooks.Open(originalFilePath, ReadOnly:=True)
    
    ' 检查原始工作簿中是否有同名sheet
    On Error Resume Next
    Set origWs = origWb.Worksheets(sheetName)
    If Err.Number <> 0 Then
        origWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "原始工作簿中未找到sheet: " & sheetName, vbExclamation, "CDM Toolkit"
        DM_TrackChanges = -1
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    ' 遍历所有有值单元格进行比较
    Dim usedRange As Range
    Set usedRange = ws.UsedRange
    
    For Each cell In usedRange
        On Error Resume Next
        origVal = origWs.Cells(cell.Row, cell.Column).Value
        On Error GoTo ErrHandler
        
        ' 值不同 → 标记修改
        If CStr(cell.Value) <> CStr(origVal) Then
            changes = changes + 1
            
            ' 清除旧批注
            If Not cell.Comment Is Nothing Then cell.Comment.Delete
            
            ' 添加批注
            If addComments Then
                Set cmt = cell.AddComment
                cmt.Text Text:="CDM Track Changes:" & vbCrLf & _
                    "旧值: " & IIf(IsEmpty(origVal), "(空)", CStr(origVal)) & vbCrLf & _
                    "新值: " & IIf(IsEmpty(cell.Value), "(空)", CStr(cell.Value))
                cmt.Shape.TextFrame.AutoSize = True
            End If
            
            ' 高亮标记
            If highlightChanges Then
                cell.Interior.Color = RGB(146, 208, 80)   ' 绿色填充
                cell.Borders.Color = RGB(255, 0, 0)         ' 红色边框
                cell.Borders.Weight = xlMedium
            End If
        End If
    Next cell
    
    origWb.Close SaveChanges:=False
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_TrackChanges = changes
    Exit Function
    
ErrHandler:
    On Error Resume Next
    origWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_TrackChanges = -1
    MsgBox "DM_TrackChanges 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_ClearTrackChanges
' 清除所有修改痕迹（批注、填充色、边框）
'============================================================================
Public Sub DM_ClearTrackChanges(ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    Dim cell As Range
    
    Application.ScreenUpdating = False
    
    For Each cell In ws.UsedRange
        If Not cell.Comment Is Nothing Then cell.Comment.Delete
        cell.Interior.Pattern = xlNone
        cell.Borders.LineStyle = xlNone
    Next cell
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_ClearTrackChanges 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_ChangesReport
' 生成结构化修改报告到新sheet
' 参数:
'   ws               — 当前工作表
'   originalFilePath — 原始工作簿文件路径
'============================================================================
Public Sub DM_ChangesReport(ByVal ws As Worksheet, ByVal originalFilePath As String)
    
    On Error GoTo ErrHandler
    
    Dim origWb As Workbook
    Dim origWs As Worksheet
    Dim report As Worksheet
    Dim cell As Range
    Dim origVal As Variant
    Dim currVal As Variant
    Dim reportRow As Long
    Dim changeType As String
    Dim sheetName As String
    
    sheetName = ws.Name
    
    If Dir(originalFilePath) = "" Then
        MsgBox "原始文件未找到: " & originalFilePath, vbExclamation, "CDM Toolkit"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Set origWb = Workbooks.Open(originalFilePath, ReadOnly:=True)
    
    On Error Resume Next
    Set origWs = origWb.Worksheets(sheetName)
    If Err.Number <> 0 Then
        origWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "原始工作簿中未找到sheet: " & sheetName, vbExclamation, "CDM Toolkit"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    ' 创建报告sheet
    On Error Resume Next
    ws.Parent.Worksheets("Changes Report").Delete
    On Error GoTo ErrHandler
    Set report = ws.Parent.Worksheets.Add(Before:=ws.Parent.Worksheets(1))
    report.Name = "Changes Report"
    
    ' 报告表头
    report.Range("A1").Value = "变更地址"
    report.Range("B1").Value = "变更类型"
    report.Range("C1").Value = "旧值"
    report.Range("D1").Value = "新值"
    report.Range("A1:D1").Font.Bold = True
    report.Range("A1:D1").Interior.Color = RGB(68, 114, 196)
    report.Range("A1:D1").Font.Color = vbWhite
    
    reportRow = 2
    
    For Each cell In ws.UsedRange
        On Error Resume Next
        origVal = origWs.Cells(cell.Row, cell.Column).Value
        On Error GoTo ErrHandler
        
        currVal = cell.Value
        
        If IsEmpty(origVal) And Not IsEmpty(currVal) Then
            changeType = "新增"
        ElseIf Not IsEmpty(origVal) And IsEmpty(currVal) Then
            changeType = "删除"
        ElseIf CStr(currVal) <> CStr(origVal) Then
            changeType = "修改"
        Else
            GoTo SkipCell
        End If
        
        report.Cells(reportRow, 1).Value = cell.Address(False, False)
        report.Cells(reportRow, 2).Value = changeType
        report.Cells(reportRow, 3).Value = IIf(IsEmpty(origVal), "(空)", CStr(origVal))
        report.Cells(reportRow, 4).Value = IIf(IsEmpty(currVal), "(空)", CStr(currVal))
        reportRow = reportRow + 1
        
SkipCell:
    Next cell
    
    report.Columns("A:D").AutoFit
    
    origWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
    MsgBox "修改报告已生成。共检测到 " & (reportRow - 2) & " 处变更。", _
        vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    origWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "DM_ChangesReport 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 11: modSpecialChars.bas — 特殊字符扫描

**Files:**
- Create: `vba/modules/modSpecialChars.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modSpecialChars"
Option Explicit

'============================================================================
' 模块: modSpecialChars
' 描述: CDM特殊字符扫描 — 扫描工作表中的非标准字符，支持EDC（Rave/Clinflash）报告
'        Special Character Scanner — find cells with non-standard characters
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

' 默认允许的字符模式（ASCII字母数字 + 常见标点和空白）
Private Const DEFAULT_ALLOWED As String = "a-zA-Z0-9\s\-_.,:;()/"

'============================================================================
' DM_ScanSpecialChars
' 扫描工作表中包含允许字符集以外字符的单元格
' 参数:
'   ws             — 目标工作表
'   allowedPattern — 自定义允许字符的VBScript正则模式，省略则使用默认值
' 返回: Scripting.Dictionary — {单元格地址: 单元格值}
' 用法:
'   Dim findings As Object
'   Set findings = DM_ScanSpecialChars(ActiveSheet)
'   If findings.count > 0 Then MsgBox "发现 " & findings.count & " 个含特殊字符的单元格"
'============================================================================
Public Function DM_ScanSpecialChars(ByVal ws As Worksheet, _
    Optional ByVal allowedPattern As String = "") As Object
    
    On Error GoTo ErrHandler
    
    Dim findings As Object
    Dim cell As Range
    Dim regex As Object
    Dim pattern As String
    
    Set findings = CreateObject("Scripting.Dictionary")
    
    ' 构建正则模式
    If Len(allowedPattern) = 0 Then
        pattern = "^[" & DEFAULT_ALLOWED & "]+$"
    Else
        pattern = "^[" & allowedPattern & "]+$"
    End If
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.Global = False
    
    Application.ScreenUpdating = False
    
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            If Not regex.Test(CStr(cell.Value)) Then
                findings.Add cell.Address(False, False), CStr(cell.Value)
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Set DM_ScanSpecialChars = findings
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_ScanSpecialChars 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_HighlightSpecialChars
' 查找并高亮包含非标准字符的单元格
' 参数:
'   ws             — 目标工作表
'   highlightColor — 高亮颜色，默认 vbRed
'   allowedPattern — 自定义允许字符模式
' 返回: Long — 高亮的单元格数量
'============================================================================
Public Function DM_HighlightSpecialChars(ByVal ws As Worksheet, _
    Optional ByVal highlightColor As Long = vbRed, _
    Optional ByVal allowedPattern As String = "") As Long
    
    On Error GoTo ErrHandler
    
    Dim findings As Object
    Dim key As Variant
    Dim cellAddr As String
    Dim count As Long
    
    Set findings = DM_ScanSpecialChars(ws, allowedPattern)
    
    Application.ScreenUpdating = False
    
    For Each key In findings.Keys
        ws.Range(CStr(key)).Interior.Color = highlightColor
        count = count + 1
    Next key
    
    Application.ScreenUpdating = True
    DM_HighlightSpecialChars = count
    Exit Function
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_HighlightSpecialChars 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

'============================================================================
' DM_EDCSpecialCharReport
' 生成EDC专用特殊字符报告（Rave 或 Clinflash 格式）
' 参数:
'   ws     — 目标工作表
'   system — EDC系统名称，"Rave" 或 "Clinflash"
' 用法:
'   DM_EDCSpecialCharReport ActiveSheet, "Rave"
'============================================================================
Public Sub DM_EDCSpecialCharReport(ByVal ws As Worksheet, _
    Optional ByVal system As String = "Rave")
    
    On Error GoTo ErrHandler
    
    Dim findings As Object
    Dim report As Worksheet
    Dim key As Variant
    Dim reportRow As Long
    Dim reportName As String
    
    Set findings = DM_ScanSpecialChars(ws)
    
    If findings.count = 0 Then
        MsgBox "未发现特殊字符。", vbInformation, "CDM Toolkit - " & system
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 创建报告sheet
    reportName = system & "_SpecialChars"
    On Error Resume Next
    ws.Parent.Worksheets(reportName).Delete
    On Error GoTo ErrHandler
    Set report = ws.Parent.Worksheets.Add
    report.Name = reportName
    
    ' 报告标题
    report.Range("A1").Value = system & " 特殊字符报告"
    report.Range("A1").Font.Bold = True
    report.Range("A1").Font.Size = 14
    
    report.Range("A2").Value = "源Sheet: " & ws.Name
    report.Range("A3").Value = "检测时间: " & Now()
    report.Range("A4").Value = "共发现: " & findings.count & " 处"
    
    ' 表头
    report.Range("A6").Value = "序号"
    report.Range("B6").Value = "单元格地址"
    report.Range("C6").Value = "内容"
    report.Range("A6:C6").Font.Bold = True
    report.Range("A6:C6").Interior.Color = RGB(68, 114, 196)
    report.Range("A6:C6").Font.Color = vbWhite
    
    reportRow = 7
    Dim idx As Long: idx = 1
    
    For Each key In findings.Keys
        report.Cells(reportRow, 1).Value = idx
        report.Cells(reportRow, 2).Value = CStr(key)
        report.Cells(reportRow, 3).Value = Left(CStr(findings(key)), 200)  ' 截断长文本
        reportRow = reportRow + 1
        idx = idx + 1
    Next key
    
    report.Columns("A:C").AutoFit
    report.Columns("C").ColumnWidth = 60
    
    ' 同时高亮原sheet
    DM_HighlightSpecialChars ws
    
    Application.ScreenUpdating = True
    MsgBox system & " 特殊字符报告已生成。共发现 " & findings.count & " 处特殊字符。", _
        vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_EDCSpecialCharReport 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证编译**

---

## 阶段五：收尾（文件目录 + Python补充）

### Task 12: modFileListing.bas — 文件目录生成

**Files:**
- Create: `vba/modules/modFileListing.bas`

- [ ] **Step 1: 创建模块文件**

```vb
Attribute VB_Name = "modFileListing"
Option Explicit

'============================================================================
' 模块: modFileListing
' 描述: CDM文件目录生成 — 生成文件夹内容列表（含超链接）和树形结构
'        File Directory Listing — generate folder contents with hyperlinks + folder tree
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: Microsoft Scripting Runtime（后期绑定，无需手动引用）
'============================================================================

'============================================================================
' DM_GenerateFileListing
' 列出文件夹中所有文件并添加超链接
' 参数:
'   ws                — 输出目标工作表
'   folderPath        — 要扫描的文件夹路径；为空则弹出选择对话框
'   includeSubFolders — 是否递归包含子文件夹，默认 True
'   fileFilter        — 文件过滤器，如 "*.xlsx" 或 "*.pdf"，默认 "*.*"
' 用法:
'   DM_GenerateFileListing ActiveSheet                       ' 弹窗选择
'   DM_GenerateFileListing ActiveSheet, "C:\Data", True, "*.xlsx"
'============================================================================
Public Sub DM_GenerateFileListing(ByVal ws As Worksheet, _
    Optional ByVal folderPath As String = "", _
    Optional ByVal includeSubFolders As Boolean = True, _
    Optional ByVal fileFilter As String = "*.*")
    
    On Error GoTo ErrHandler
    
    ' 选择文件夹
    If Len(folderPath) = 0 Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "选择要生成文件列表的文件夹"
            If .Show = -1 Then
                folderPath = .SelectedItems(1)
            Else
                Exit Sub
            End If
        End With
    End If
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' 准备输出
    Application.ScreenUpdating = False
    ws.Cells.Clear
    
    ' 标题行
    ws.Range("A1").Value = "文件夹路径"
    ws.Range("B1").Value = "文件名"
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = RGB(68, 114, 196)
    ws.Range("A1:B1").Font.Color = vbWhite
    
    ' 递归扫描
    DM_ScanFolder ws, folderPath, includeSubFolders, fileFilter
    
    ws.Columns("A:B").AutoFit
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_GenerateFileListing 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

' 递归扫描文件夹
Private Sub DM_ScanFolder(ByVal ws As Worksheet, _
    ByVal folderPath As String, _
    ByVal includeSubFolders As Boolean, _
    ByVal fileFilter As String)
    
    On Error GoTo ErrHandler
    
    Dim fso As Object, folder As Object, file As Object, subFolder As Object
    Dim lastRow As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' 列出文件
    For Each file In folder.Files
        ' 检查文件过滤器
        If DM_MatchFilter(file.Name, fileFilter) Then
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
            ws.Cells(lastRow, 1).Value = file.ParentFolder.Path
            ws.Cells(lastRow, 2).Value = file.Name
            ws.Hyperlinks.Add Anchor:=ws.Cells(lastRow, 2), _
                Address:=file.Path, ScreenTip:=file.Path
        End If
    Next file
    
    ' 递归子文件夹
    If includeSubFolders Then
        For Each subFolder In folder.SubFolders
            DM_ScanFolder ws, subFolder.Path, includeSubFolders, fileFilter
        Next subFolder
    End If
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    Exit Sub
    
ErrHandler:
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

' 简单通配符匹配
Private Function DM_MatchFilter(ByVal fileName As String, ByVal filter As String) As Boolean
    If filter = "*.*" Or filter = "*" Then
        DM_MatchFilter = True
    ElseIf Left(filter, 2) = "*." Then
        Dim ext As String
        ext = Mid(filter, 2)  ' ".xlsx", ".pdf" 等
        DM_MatchFilter = (LCase(Right(fileName, Len(ext))) = LCase(ext))
    Else
        DM_MatchFilter = (LCase(fileName) Like LCase(filter))
    End If
End Function

'============================================================================
' DM_GenerateFolderTree
' 递归树形结构展示文件夹（带缩进和颜色标记）
' 参数:
'   ws         — 输出目标工作表
'   folderPath — 要扫描的文件夹路径；为空则弹出选择对话框
' 用法:
'   DM_GenerateFolderTree ActiveSheet
'============================================================================
Public Sub DM_GenerateFolderTree(ByVal ws As Worksheet, _
    Optional ByVal folderPath As String = "")
    
    On Error GoTo ErrHandler
    
    If Len(folderPath) = 0 Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "选择要生成树形目录的文件夹"
            If .Show = -1 Then
                folderPath = .SelectedItems(1)
            Else
                Exit Sub
            End If
        End With
    End If
    
    Dim fso As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    Application.ScreenUpdating = False
    ws.Cells.Clear
    
    ws.Cells(1, 1).Value = "文件夹树形结构"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    
    ' 根文件夹
    ws.Cells(2, 1).Value = folder.Path
    ws.Cells(2, 1).Interior.Color = RGB(68, 114, 196)
    ws.Cells(2, 1).Font.Color = vbWhite
    ws.Cells(2, 1).Font.Bold = True
    
    ' 递归构建树
    DM_BuildTree ws, folder, 3, 1, True
    
    ws.Columns("A").AutoFit
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_GenerateFolderTree 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

' 递归构建树形结构
Private Sub DM_BuildTree(ByVal ws As Worksheet, _
    ByVal folder As Object, _
    ByRef rowNum As Long, _
    ByVal depth As Long, _
    ByVal isLastItem As Boolean)
    
    On Error Resume Next
    
    Dim subFolder As Object, file As Object
    Dim indent As String
    Dim prefix As String
    
    indent = String(depth * 4, " ")
    
    ' 子文件夹
    For Each subFolder In folder.SubFolders
        ws.Cells(rowNum, 1).Value = indent & "|- " & subFolder.Name & "\"
        ws.Cells(rowNum, 1).Interior.Color = vbYellow
        rowNum = rowNum + 1
        DM_BuildTree ws, subFolder, rowNum, depth + 1, False
    Next subFolder
    
    ' 文件
    For Each file In folder.Files
        ws.Cells(rowNum, 1).Value = indent & "|  " & file.Name
        ws.Hyperlinks.Add Anchor:=ws.Cells(rowNum, 1), _
            Address:=file.Path, ScreenTip:=file.Path
        rowNum = rowNum + 1
    Next file
    
    On Error GoTo 0
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 13: cdm_toolkit_plus.py — Python 补充

**Files:**
- Create: `cdm_toolkit_plus.py`

- [ ] **Step 1: 创建Python补充文件**

```python
"""
CDM Excel Toolkit Plus — Python supplement adding features discovered
during VBA analysis that are missing from cdm_toolkit.py.

New features:
  1. File directory listing → Excel with hyperlinks
  2. Folder tree → Excel with indentation
  3. Track changes with real Excel comments (via COM/xlwings)
  4. Export .bas module into .xlsm workbook
"""

from __future__ import annotations
import os
import re
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink


# ============================================================
#  Style constants (match cdm_toolkit.py)
# ============================================================
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TOC_FONT = Font(color="0563C1", underline="single", size=11)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GREEN_FILL = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
RED_BORDER = Border(
    left=Side(style="medium", color="FF0000"),
    right=Side(style="medium", color="FF0000"),
    top=Side(style="medium", color="FF0000"),
    bottom=Side(style="medium", color="FF0000"),
)


# ============================================================
#  1. File Directory Listing → Excel
# ============================================================
def generate_file_listing(
    folder_path: str,
    output_path: str,
    include_subfolders: bool = True,
    file_pattern: str = "*.*",
) -> Workbook:
    """
    Generate an Excel file listing all files in a folder with hyperlinks.

    Parameters
    ----------
    folder_path : str
        Root folder to scan.
    output_path : str
        Where to save the Excel file.
    include_subfolders : bool
        Recursively scan subfolders (default True).
    file_pattern : str
        File filter, e.g. "*.xlsx" or "*.pdf" (default "*.*").

    Returns
    -------
    Workbook
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "File Listing"

    # Header
    ws.cell(row=1, column=1, value="Folder Path").font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL
    ws.cell(row=1, column=2, value="File Name").font = HEADER_FONT
    ws.cell(row=1, column=2).fill = HEADER_FILL

    # Parse file pattern
    if file_pattern in ("*.*", "*"):
        ext_filter = None
    elif file_pattern.startswith("*."):
        ext_filter = file_pattern[1:].lower()  # ".xlsx"
    else:
        ext_filter = file_pattern.lower()

    row_idx = 2

    def _scan(path: str) -> None:
        nonlocal row_idx
        try:
            for entry in os.scandir(path):
                if entry.is_file():
                    name = entry.name
                    if ext_filter is None or name.lower().endswith(ext_filter):
                        ws.cell(row=row_idx, column=1, value=os.path.dirname(entry.path))
                        cell = ws.cell(row=row_idx, column=2, value=name)
                        cell.font = TOC_FONT
                        cell.hyperlink = Hyperlink(
                            ref=cell.coordinate,
                            target=entry.path,
                            display=name,
                        )
                        row_idx += 1
                elif entry.is_dir() and include_subfolders:
                    _scan(entry.path)
        except PermissionError:
            pass  # Skip inaccessible folders

    _scan(folder_path)

    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 50
    wb.save(output_path)
    return wb


# ============================================================
#  2. Folder Tree → Excel
# ============================================================
def generate_folder_tree(
    folder_path: str,
    output_path: str,
) -> Workbook:
    """
    Generate a recursive tree-view folder structure as Excel.

    Parameters
    ----------
    folder_path : str
        Root folder to scan.
    output_path : str
        Where to save the Excel file.

    Returns
    -------
    Workbook
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Folder Tree"

    ws.cell(row=1, column=1, value="Folder Tree Structure").font = Font(bold=True, size=14)
    ws.cell(row=2, column=1, value=os.path.abspath(folder_path))
    ws.cell(row=2, column=1).fill = HEADER_FILL
    ws.cell(row=2, column=1).font = HEADER_FONT

    row_idx = 3

    def _build_tree(path: str, depth: int = 0) -> None:
        nonlocal row_idx
        indent = "    " * depth
        try:
            entries = sorted(os.scandir(path), key=lambda e: (not e.is_dir(), e.name.lower()))
            for entry in entries:
                if entry.is_dir():
                    ws.cell(row=row_idx, column=1, value=f"{indent}|- {entry.name}\\")
                    ws.cell(row=row_idx, column=1).fill = YELLOW_FILL
                    row_idx += 1
                    _build_tree(entry.path, depth + 1)
                else:
                    cell = ws.cell(row=row_idx, column=1, value=f"{indent}|  {entry.name}")
                    cell.font = TOC_FONT
                    cell.hyperlink = Hyperlink(
                        ref=cell.coordinate,
                        target=entry.path,
                        display=entry.name,
                    )
                    row_idx += 1
        except PermissionError:
            ws.cell(row=row_idx, column=1, value=f"{indent}[Access Denied]")
            row_idx += 1

    _build_tree(folder_path)
    ws.column_dimensions["A"].width = 80
    wb.save(output_path)
    return wb


# ============================================================
#  3. Track Changes with Excel Comments (xlwings)
# ============================================================
def track_changes_with_comments(
    wb: Workbook,
    original_path: str,
) -> int:
    """
    Track changes by comparing a workbook against an original copy.
    Tries to use xlwings for real Excel comments; falls back to color marking.

    Parameters
    ----------
    wb : Workbook (openpyxl)
        Current workbook (may be unsaved).
    original_path : str
        Path to the original workbook for comparison.

    Returns
    -------
    int — number of changed cells detected.
    """
    import openpyxl

    orig_wb = openpyxl.load_workbook(original_path, data_only=True)
    changes = 0

    for ws in wb.worksheets:
        if ws.title not in orig_wb.sheetnames:
            continue
        orig_ws = orig_wb[ws.title]

        for row in ws.iter_rows():
            for cell in row:
                try:
                    orig_val = orig_ws.cell(row=cell.row, column=cell.column).value
                except (KeyError, ValueError):
                    orig_val = None

                if str(cell.value) != str(orig_val):
                    # Mark with green fill + red border (openpyxl comment support is limited)
                    cell.fill = GREEN_FILL
                    cell.border = RED_BORDER
                    # Store old→new in a hidden cell attribute for reference
                    cell._comment_old = orig_val
                    changes += 1

    orig_wb.close()
    return changes


# ============================================================
#  4. Export .bas module into .xlsm workbook (Windows only, via COM)
# ============================================================
def export_vba_module(xlsm_path: str, bas_path: str) -> bool:
    """
    Embed a .bas VBA module into an existing .xlsm workbook using COM.

    Parameters
    ----------
    xlsm_path : str
        Path to the .xlsm workbook.
    bas_path : str
        Path to the .bas file to import.

    Returns
    -------
    bool — True if successful, False otherwise.
    """
    if not os.path.exists(bas_path):
        print(f"Error: .bas file not found: {bas_path}")
        return False

    # Ensure output is .xlsm
    if not xlsm_path.lower().endswith(".xlsm"):
        print("Warning: output should be .xlsm for macro support. Renaming.")
        xlsm_path = xlsm_path.rsplit(".", 1)[0] + ".xlsm"

    # Create a fresh .xlsm if it doesn't exist
    if not os.path.exists(xlsm_path):
        wb = Workbook()
        wb.save(xlsm_path)

    # Attempt COM import
    try:
        import win32com.client  # type: ignore

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb_com = excel.Workbooks.Open(os.path.abspath(xlsm_path))
        module_name = os.path.splitext(os.path.basename(bas_path))[0]

        # Remove existing module if present
        try:
            wb_com.VBProject.VBComponents(module_name).Activate
            wb_com.VBProject.VBComponents.Remove(
                wb_com.VBProject.VBComponents(module_name)
            )
        except Exception:
            pass

        # Import .bas file
        wb_com.VBProject.VBComponents.Import(os.path.abspath(bas_path))
        wb_com.Save()
        wb_com.Close()
        excel.Quit()

        print(f"Successfully imported {os.path.basename(bas_path)} into {xlsm_path}")
        return True

    except ImportError:
        print("pywin32 not installed. Install with: pip install pywin32")
        print(f"Manually import {bas_path} into {xlsm_path} via Excel VBA Editor.")
        return False
    except Exception as e:
        print(f"COM import failed: {e}")
        print("Ensure Excel 'Trust access to the VBA project object model' is enabled.")
        print(f"Manually import {bas_path} into {xlsm_path} via Excel VBA Editor.")
        return False


# ============================================================
#  Main entry point
# ============================================================
if __name__ == "__main__":
    print("CDM Excel Toolkit Plus — use from your own scripts:")
    print("  from cdm_toolkit_plus import *")
    print("  generate_file_listing('C:/Data', 'listing.xlsx')")
    print("  generate_folder_tree('C:/Data', 'tree.xlsx')")
    print("  track_changes_with_comments(wb, 'original.xlsx')")
    print("  export_vba_module('workbook.xlsm', 'module.bas')")
```

- [ ] **Step 2: 运行Python导入测试**

```bash
cd e:/Project/Tools/Tools_DM && python -c "from cdm_toolkit_plus import *; print('Import OK')"
```

---

## 阶段六：集成（主模块 + 安装 + 文档 + 测试）

### Task 14: CDM_Toolkit.bas — 主汇总模块

**Files:**
- Create: `vba/CDM_Toolkit.bas`

- [ ] **Step 1: 创建主汇总模块**

```vb
Attribute VB_Name = "CDM_Toolkit"
Option Explicit

'============================================================================
' 模块: CDM_Toolkit
' 描述: CDM Toolkit 主模块 — 提供版本信息和快捷调用入口
'        Master module — version info and quick-launch entry points
' 版本: 1.0 / 2026-05-30
'
' 包含的子模块（按功能分组）:
'   [数据清理] modTextStandardise, modHighlightDup, modBlankRows
'   [文档管理] modTOC, modSheetSplit, modWorkbookMerge
'   [工作表]   modSheetManager
'   [审核追溯] modTrackChanges, modSpecialChars
'   [计算转换] modDateCalc, modUnitConvert
'   [工具]     modFileListing
'
' 安装方法:
'   1. 在Excel中按 Alt+F11 打开VBA编辑器
'   2. File → Import File → 选择 vba/modules/ 下所有 .bas 文件
'   3. 再导入本文件 (CDM_Toolkit.bas)
'   或运行 CDM_Install.bas 中的 DM_InstallAll 一键安装
'============================================================================

Public Const CDM_TOOLKIT_VERSION As String = "1.0"
Public Const CDM_TOOLKIT_DATE As String = "2026-05-30"

'============================================================================
' DM_About
' 显示CDM Toolkit版本信息
'============================================================================
Public Sub DM_About()
    MsgBox "CDM Excel Toolkit v" & CDM_TOOLKIT_VERSION & vbCrLf & vbCrLf & _
        "临床数据管理 (CDM) Excel 宏工具包" & vbCrLf & _
        "Clinical Data Management Excel Macro Toolkit" & vbCrLf & vbCrLf & _
        "包含 12 个功能模块:" & vbCrLf & _
        "  [数据清理] 文本标准化、重复高亮、空白行管理" & vbCrLf & _
        "  [文档管理] 目录生成、Sheet拆分、工作簿合并" & vbCrLf & _
        "  [工作表]   Sheet管理（重命名/列表/删除）" & vbCrLf & _
        "  [审核追溯] 修改痕迹、特殊字符扫描" & vbCrLf & _
        "  [计算转换] 日期计算、单位转换" & vbCrLf & _
        "  [工具]     文件目录生成" & vbCrLf & vbCrLf & _
        "日期: " & CDM_TOOLKIT_DATE, _
        vbInformation, "CDM Toolkit"
End Sub

'============================================================================
' DM_QuickClean
' 一键执行标准数据清理流程（术语标准化 + 重复高亮 + 删除空白行）
' 参数:
'   ws — 目标工作表
'============================================================================
Public Sub DM_QuickClean(ByVal ws As Worksheet)
    On Error GoTo ErrHandler
    
    Dim count1 As Long, count2 As Long, count3 As Long
    
    Application.StatusBar = "CDM Toolkit: 正在标准化术语..."
    count1 = DM_TextStandardise(ws)
    
    Application.StatusBar = "CDM Toolkit: 正在高亮重复值..."
    DM_HighlightDuplicates ws
    
    Application.StatusBar = "CDM Toolkit: 正在删除空白行..."
    count3 = DM_DeleteBlankRows(ws)
    
    Application.StatusBar = False
    MsgBox "快速清理完成!" & vbCrLf & _
        "  术语替换: " & count1 & " 处" & vbCrLf & _
        "  空白行删除: " & count3 & " 行", _
        vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    Application.StatusBar = False
    MsgBox "DM_QuickClean 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_QuickDoc
' 一键执行标准文档处理流程（目录生成 + 日期计算）
' 参数:
'   wb — 目标工作簿
'============================================================================
Public Sub DM_QuickDoc(ByVal wb As Workbook)
    On Error GoTo ErrHandler
    
    Application.StatusBar = "CDM Toolkit: 正在生成TOC..."
    DM_GenerateTOC wb
    
    Application.StatusBar = False
    MsgBox "文档处理完成! 已生成TOC目录。", vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    Application.StatusBar = False
    MsgBox "DM_QuickDoc 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 15: CDM_Install.bas — 一键安装模块

**Files:**
- Create: `vba/CDM_Install.bas`

- [ ] **Step 1: 创建安装模块**

```vb
Attribute VB_Name = "CDM_Install"
Option Explicit

'============================================================================
' 模块: CDM_Install
' 描述: CDM Toolkit 一键安装 — 将所有CDM宏注册到 Personal.xlsb
'        One-click install — register all CDM macros to Personal.xlsb
' 版本: 1.0 / 2026-05-30
' 依赖: 需要先导入 vba/modules/ 下所有模块
'
' 安装步骤:
'   1. 打开Excel，按 Alt+F11 进入VBA编辑器
'   2. File → Import File → 选择并导入 vba/modules/ 下全部 .bas 文件
'   3. 再导入本文件 (CDM_Install.bas)
'   4. 运行 DM_InstallAll
'   5. 所有宏将被复制到 Personal.xlsb，之后任意Excel文件都可以使用
'============================================================================

Private Const MODULE_NAMES As String = _
    "modTextStandardise,modHighlightDup,modBlankRows," & _
    "modTOC,modSheetSplit,modWorkbookMerge," & _
    "modSheetManager,modTrackChanges," & _
    "modDateCalc,modUnitConvert,modSpecialChars,modFileListing"

'============================================================================
' DM_InstallAll
' 将所有CDM模块复制到Personal.xlsb（个人宏工作簿）
'============================================================================
Public Sub DM_InstallAll()
    On Error GoTo ErrHandler
    
    Dim personalWb As Workbook
    Dim personalPath As String
    Dim modules() As String
    Dim i As Long
    Dim comp As Object ' VBComponent
    Dim srcComp As Object
    Dim installed As Long
    Dim failed As String
    
    modules = Split(MODULE_NAMES, ",")
    
    ' 确保 Personal.xlsb 存在
    personalPath = Application.StartupPath & "\PERSONAL.XLSB"
    
    On Error Resume Next
    Set personalWb = Workbooks("PERSONAL.XLSB")
    If Err.Number <> 0 Then
        ' Personal.xlsb 不存在，先创建一个（录制一个宏即可触发创建）
        MsgBox "未找到 Personal.xlsb。正在创建..." & vbCrLf & _
            "请先运行一次 'DM_CreatePersonal' 创建个人宏工作簿。", _
            vbInformation, "CDM Toolkit"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    For i = 0 To UBound(modules)
        Dim modName As String
        modName = Trim(modules(i))
        
        ' 检查模块是否已存在
        On Error Resume Next
        Set srcComp = ThisWorkbook.VBProject.VBComponents(modName)
        If Err.Number <> 0 Then
            failed = failed & modName & " (未在当前工作簿中找到)" & vbCrLf
            Err.Clear
            GoTo NextModule
        End If
        On Error GoTo ErrHandler
        
        ' 删除Personal中的旧版本
        On Error Resume Next
        personalWb.VBProject.VBComponents(modName).Activate
        If Err.Number = 0 Then
            personalWb.VBProject.VBComponents.Remove _
                personalWb.VBProject.VBComponents(modName)
        End If
        Err.Clear
        On Error GoTo ErrHandler
        
        ' 导出→导入方式复制模块
        Dim tempPath As String
        tempPath = Environ("TEMP") & "\" & modName & ".bas"
        srcComp.Export tempPath
        personalWb.VBProject.VBComponents.Import tempPath
        Kill tempPath
        
        installed = installed + 1
        
NextModule:
    Next i
    
    personalWb.Save
    
    Application.ScreenUpdating = True
    
    Dim msg As String
    msg = "CDM Toolkit 安装完成!" & vbCrLf & _
        "成功安装: " & installed & " / " & (UBound(modules) + 1) & " 个模块" & vbCrLf
    If Len(failed) > 0 Then
        msg = msg & vbCrLf & "未安装的模块:" & vbCrLf & failed
    End If
    MsgBox msg, vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_InstallAll 出错: " & Err.Description & vbCrLf & vbCrLf & _
        "请确保已启用 '信任对VBA工程对象模型的访问':" & vbCrLf & _
        "Excel → 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置", _
        vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_CreatePersonal
' 创建 Personal.xlsb 个人宏工作簿
'============================================================================
Public Sub DM_CreatePersonal()
    On Error Resume Next
    
    ' 录制一个简单宏即可触发Excel创建Personal.xlsb
    ' 方法：开始录制宏 → 选择存储在"个人宏工作簿" → 停止录制
    MsgBox "请按以下步骤创建 Personal.xlsb:" & vbCrLf & vbCrLf & _
        "1. 点击 '视图' → '宏' → '录制宏'" & vbCrLf & _
        "2. '保存在' 选择 '个人宏工作簿'" & vbCrLf & _
        "3. 点击 '确定'，然后立即点击 '停止录制'" & vbCrLf & vbCrLf & _
        "创建完成后，再次运行 DM_InstallAll。", _
        vbInformation, "CDM Toolkit"
End Sub

'============================================================================
' DM_UninstallAll
' 从 Personal.xlsb 中移除所有CDM模块
'============================================================================
Public Sub DM_UninstallAll()
    On Error GoTo ErrHandler
    
    Dim personalWb As Workbook
    Dim modules() As String
    Dim i As Long
    Dim removed As Long
    
    On Error Resume Next
    Set personalWb = Workbooks("PERSONAL.XLSB")
    If Err.Number <> 0 Then
        MsgBox "未找到 Personal.xlsb。", vbInformation, "CDM Toolkit"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    modules = Split(MODULE_NAMES, ",")
    
    For i = 0 To UBound(modules)
        Dim modName As String
        modName = Trim(modules(i))
        
        On Error Resume Next
        personalWb.VBProject.VBComponents(modName).Activate
        If Err.Number = 0 Then
            personalWb.VBProject.VBComponents.Remove _
                personalWb.VBProject.VBComponents(modName)
            removed = removed + 1
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next i
    
    ' 也移除主模块
    On Error Resume Next
    personalWb.VBProject.VBComponents("CDM_Toolkit").Activate
    If Err.Number = 0 Then
        personalWb.VBProject.VBComponents.Remove _
            personalWb.VBProject.VBComponents("CDM_Toolkit")
    End If
    personalWb.VBProject.VBComponents("CDM_Install").Activate
    If Err.Number = 0 Then
        personalWb.VBProject.VBComponents.Remove _
            personalWb.VBProject.VBComponents("CDM_Install")
    End If
    Err.Clear
    On Error GoTo ErrHandler
    
    personalWb.Save
    MsgBox "已卸载 " & removed & " 个CDM模块。" & vbCrLf & _
        "CDM_Toolkit 和 CDM_Install 模块也已移除。", vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    MsgBox "DM_UninstallAll 出错: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 2: 验证编译**

---

### Task 16: vba/README.md — VBA 使用说明书

**Files:**
- Create: `vba/README.md`

- [ ] **Step 1: 创建使用说明书**

```markdown
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
| `DM_TextStandardise` | 批量标准化CDM术语（Recovered→Recovered/Resolved 等） | `DM_TextStandardise ActiveSheet` |
| `DM_HighlightDuplicates` | 高亮连续重复的单元格（黄色） | `DM_HighlightDuplicates ActiveSheet` |
| `DM_InsertBlankEveryN` | 每N行插入空白行 | `DM_InsertBlankEveryN ActiveSheet, 5` |
| `DM_InsertBlankBetweenGroups` | 组间插入空白行 | `DM_InsertBlankBetweenGroups ActiveSheet, "A"` |
| `DM_DeleteBlankRows` | 删除完全空白的行 | `DM_DeleteBlankRows ActiveSheet` |

### 文档管理

| 宏名 | 功能 | 用法示例 |
|------|------|---------|
| `DM_GenerateTOC` | 生成带超链接的目录sheet | `DM_GenerateTOC ActiveWorkbook` |
| `DM_GenerateDeriveTOC` | SAS Derive风格目录（B列=描述） | `DM_GenerateDeriveTOC ActiveWorkbook, "V2"` |
| `DM_SplitSheetByColumn` | 按列值拆分sheet | `DM_SplitSheetByColumn ActiveSheet, "A"` |
| `DM_MergeWorkbooksDialog` | 合并多个工作簿（弹窗选择） | `DM_MergeWorkbooksDialog` |

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
- 安装到 Personal.xlsb 需要启用"信任对VBA工程对象模型的访问"
```

- [ ] **Step 2: 完成（创建即可）**

---

### Task 17: 测试文件

**Files:**
- Create: `tests/test_cdm_plus.py`

- [ ] **Step 1: 创建Python测试**

```python
"""
Test script for cdm_toolkit_plus.py.
Creates test data and runs all new functions.
"""

import os
import sys
import shutil
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from cdm_toolkit_plus import *

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_plus_output")


def setup():
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR)
    # Create test folder structure
    test_dir = os.path.join(OUTPUT_DIR, "test_folder")
    os.makedirs(os.path.join(test_dir, "sub_a"))
    os.makedirs(os.path.join(test_dir, "sub_b"))
    Path(os.path.join(test_dir, "file1.xlsx")).touch()
    Path(os.path.join(test_dir, "file2.pdf")).touch()
    Path(os.path.join(test_dir, "sub_a", "file3.txt")).touch()
    Path(os.path.join(test_dir, "sub_b", "file4.xlsx")).touch()
    return test_dir


def test_generate_file_listing():
    print("[TEST] File listing ...", end=" ")
    test_dir = setup()
    out_path = os.path.join(OUTPUT_DIR, "file_listing.xlsx")
    wb = generate_file_listing(test_dir, out_path)
    ws = wb["File Listing"]
    # Should have header + at least 4 files
    assert ws.max_row >= 5, f"Expected >=5 rows, got {ws.max_row}"
    print("OK")


def test_generate_folder_tree():
    print("[TEST] Folder tree ...", end=" ")
    test_dir = setup()
    out_path = os.path.join(OUTPUT_DIR, "folder_tree.xlsx")
    wb = generate_folder_tree(test_dir, out_path)
    assert "Folder Tree" in wb.sheetnames
    print("OK")


def test_track_changes():
    print("[TEST] Track changes (openpyxl fallback) ...", end=" ")
    from openpyxl import Workbook
    
    # Create original
    orig_path = os.path.join(OUTPUT_DIR, "original.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Test"
    ws.cell(row=1, column=1, value="Hello")
    ws.cell(row=2, column=1, value="World")
    wb.save(orig_path)
    
    # Create modified
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "Test"
    ws2.cell(row=1, column=1, value="Hello")
    ws2.cell(row=2, column=1, value="Changed")  # Changed!
    
    changes = track_changes_with_comments(wb2, orig_path)
    assert changes >= 1, f"Expected >=1 change, got {changes}"
    print(f"OK ({changes} changes)")


if __name__ == "__main__":
    setup()
    tests = [test_generate_file_listing, test_generate_folder_tree, test_track_changes]
    passed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except Exception as e:
            print(f"FAIL: {e}")
    print(f"\nResults: {passed}/{len(tests)} passed")
```

- [ ] **Step 2: 运行测试**

```bash
cd e:/Project/Tools/Tools_DM && python tests/test_cdm_plus.py
```

---

## 实现检查清单

| # | 任务 | 文件 | 状态 |
|---|------|------|------|
| 1 | modTextStandardise | vba/modules/modTextStandardise.bas | - [ ] |
| 2 | modHighlightDup | vba/modules/modHighlightDup.bas | - [ ] |
| 3 | modBlankRows | vba/modules/modBlankRows.bas | - [ ] |
| 4 | modTOC | vba/modules/modTOC.bas | - [ ] |
| 5 | modSheetSplit | vba/modules/modSheetSplit.bas | - [ ] |
| 6 | modWorkbookMerge | vba/modules/modWorkbookMerge.bas | - [ ] |
| 7 | modSheetManager | vba/modules/modSheetManager.bas | - [ ] |
| 8 | modDateCalc | vba/modules/modDateCalc.bas | - [ ] |
| 9 | modUnitConvert | vba/modules/modUnitConvert.bas | - [ ] |
| 10 | modTrackChanges | vba/modules/modTrackChanges.bas | - [ ] |
| 11 | modSpecialChars | vba/modules/modSpecialChars.bas | - [ ] |
| 12 | modFileListing | vba/modules/modFileListing.bas | - [ ] |
| 13 | cdm_toolkit_plus.py | cdm_toolkit_plus.py | - [ ] |
| 14 | CDM_Toolkit.bas | vba/CDM_Toolkit.bas | - [ ] |
| 15 | CDM_Install.bas | vba/CDM_Install.bas | - [ ] |
| 16 | README | vba/README.md | - [ ] |
| 17 | Python测试 | tests/test_cdm_plus.py | - [ ] |
