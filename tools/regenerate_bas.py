#!/usr/bin/env python3
"""
Regenerate all .bas files with GBK encoding for Chinese Windows Excel compatibility.
Each file gets GBK-encoded Chinese comments while keeping VBA code intact.
"""
import os

VBA_DIR = r"e:\Project\Tools\Tools_DM\vba"
MOD_DIR = os.path.join(VBA_DIR, "modules")
os.makedirs(MOD_DIR, exist_ok=True)

# ================================================================
# Module 1: modTextStandardise
# ================================================================
MOD1 = """Attribute VB_Name = "modTextStandardise"
Option Explicit

'============================================================================
' 模块: modTextStandardise
' 描述: CDM文本术语标准化 — 批量查找替换非标准术语为标准CDM术语
'       Text/Term Standardisation — batch replace non-standard terms
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无（使用后期绑定 Scripting.Dictionary）
'============================================================================

Private mTermMap As Object

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

Public Function DM_GetDefaultTermMap() As Object
    Call DM_InitTermMap
    Set DM_GetDefaultTermMap = mTermMap
End Function

Public Function DM_TextStandardise(ByVal ws As Worksheet, _
    Optional ByVal termMap As Object = Nothing) As Long
    On Error GoTo ErrHandler
    Dim cell As Range, key As Variant, count As Long, map As Object
    If termMap Is Nothing Then
        Call DM_InitTermMap: Set map = mTermMap
    Else
        Set map = termMap
    End If
    If map.Count = 0 Then DM_TextStandardise = 0: Exit Function
    Application.ScreenUpdating = False
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
    MsgBox "DM_TextStandardise Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Sub DM_TextReplace(ByVal ws As Worksheet, _
    ByVal oldText As String, ByVal newText As String, _
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
    MsgBox "DM_TextReplace Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
"""

# ================================================================
# Module 2: modHighlightDup
# ================================================================
MOD2 = """Attribute VB_Name = "modHighlightDup"
Option Explicit

'============================================================================
' 模块: modHighlightDup
' 描述: CDM重复值高亮 — 标记相邻行中相同列值连续重复的单元格
'       Duplicate Highlighting — highlight consecutive duplicate cells
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

Public Sub DM_HighlightDuplicates(ByVal ws As Worksheet, _
    Optional ByVal colStart As Long = 7, _
    Optional ByVal colEnd As Long = 11, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal highlightColor As Long = 65535)
    On Error GoTo ErrHandler
    Dim lastRow As Long, colIdx As Long, rowIdx As Long
    Dim currVal As Variant, prevVal As Variant
    If colStart < 1 Or colEnd < colStart Or rowStart < 1 Then
        MsgBox "DM_HighlightDuplicates: Invalid parameters.", vbExclamation, "CDM Toolkit"
        Exit Sub
    End If
    lastRow = ws.Cells(ws.Rows.Count, colStart).End(xlUp).Row
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
    MsgBox "DM_HighlightDuplicates Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_HighlightDupSingleCol(ByVal ws As Worksheet, _
    Optional ByVal col As Long = 5, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal highlightColor As Long = 65535)
    Call DM_HighlightDuplicates(ws, col, col, rowStart, highlightColor)
End Sub
"""

# ================================================================
# Module 3: modBlankRows
# ================================================================
MOD3 = """Attribute VB_Name = "modBlankRows"
Option Explicit

'============================================================================
' 模块: modBlankRows
' 描述: CDM空白行管理 — 按规则插入/删除空白行
'       Blank Row Management — insert blank rows or delete empties
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

Public Function DM_InsertBlankEveryN(ByVal ws As Worksheet, _
    Optional ByVal n As Long = 5, _
    Optional ByVal colLetter As String = "A") As Long
    On Error GoTo ErrHandler
    Dim lastRow As Long, i As Long, inserted As Long
    If n < 1 Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
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
    MsgBox "DM_InsertBlankEveryN Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Function DM_InsertBlankBetweenGroups(ByVal ws As Worksheet, _
    Optional ByVal colLetter As String = "A") As Long
    On Error GoTo ErrHandler
    Dim lastRow As Long, i As Long, inserted As Long
    Dim currVal As Variant, nextVal As Variant
    lastRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
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
    MsgBox "DM_InsertBlankBetweenGroups Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Function DM_DeleteBlankRows(ByVal ws As Worksheet) As Long
    On Error GoTo ErrHandler
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim isEmptyRow As Boolean, deleted As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow <= 1 Then Exit Function
    If lastCol < 1 Then lastCol = 1
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For i = lastRow To 1 Step -1
        isEmptyRow = True
        For j = 1 To lastCol
            If Not IsEmpty(ws.Cells(i, j).Value) Then
                isEmptyRow = False: Exit For
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
    MsgBox "DM_DeleteBlankRows Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function
"""

# ================================================================
# Module 4: modTOC
# ================================================================
MOD4 = """Attribute VB_Name = "modTOC"
Option Explicit

'============================================================================
' 模块: modTOC
' 描述: CDM目录生成器 — 生成带超链接的工作簿目录（标准版 + SAS Derive版）
'       Table of Contents generator with hyperlinks + back-arrow navigation
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

Private Const BACK_ARROW_NAME As String = "CDM_BackToTOC"

Private Function DM_SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    DM_SheetExists = (Err.Number = 0)
    Err.Clear
End Function

Public Sub DM_GenerateTOC(ByVal wb As Workbook, _
    Optional ByVal tocName As String = "TOC", _
    Optional ByVal excludeSheets As String = "")
    On Error GoTo ErrHandler
    Dim toc As Worksheet, ws As Worksheet, excludeArr() As String
    Dim rowIdx As Long, cell As Range, shp As Shape
    Dim isExcluded As Boolean, i As Long

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    If Len(excludeSheets) > 0 Then excludeArr = Split(excludeSheets, ",")

    If DM_SheetExists(wb, tocName) Then wb.Worksheets(tocName).Delete

    Set toc = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    toc.Name = tocName
    toc.Range("A1").Value = "Table of Contents"
    toc.Range("A1").Font.Bold = True
    toc.Range("A1").Font.Size = 14
    toc.Range("A1").Interior.Color = RGB(68, 114, 196)
    toc.Range("A1").Font.Color = vbWhite

    rowIdx = 2
    For Each ws In wb.Worksheets
        If ws.Name <> tocName Then
            isExcluded = False
            If Len(excludeSheets) > 0 Then
                For i = 0 To UBound(excludeArr)
                    If Trim(ws.Name) = Trim(excludeArr(i)) Then
                        isExcluded = True: Exit For
                    End If
                Next i
            End If
            If Not isExcluded Then
                Set cell = toc.Cells(rowIdx, 1)
                cell.Value = ws.Name
                cell.Font.Color = RGB(5, 99, 193)
                cell.Font.Underline = xlUnderlineStyleSingle
                toc.Hyperlinks.Add Anchor:=cell, Address:="", _
                    SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
                On Error Resume Next
                ws.Shapes(BACK_ARROW_NAME).Delete
                On Error GoTo ErrHandler
                Set shp = ws.Shapes.AddShape(msoShapeLeftArrow, 0, 0, 20, 10)
                shp.Name = BACK_ARROW_NAME
                shp.TextFrame2.TextRange.Text = "<-"
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
    MsgBox "DM_GenerateTOC Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_GenerateDeriveTOC(ByVal wb As Workbook, _
    Optional ByVal descCell As String = "V2", _
    Optional ByVal tocName As String = "TOC")
    On Error GoTo ErrHandler
    Dim toc As Worksheet, ws As Worksheet, rowIdx As Long
    Dim cellA As Range, cellB As Range, desc As Variant

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    If DM_SheetExists(wb, tocName) Then wb.Worksheets(tocName).Delete

    Set toc = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    toc.Name = tocName

    toc.Range("A1").Value = "TOC"
    toc.Range("A1").Font.Bold = True
    toc.Range("A1").Font.Size = 14
    toc.Range("B1").Value = "Description"
    toc.Range("B1").Font.Bold = True
    toc.Range("B1").Font.Size = 14

    rowIdx = 2
    For Each ws In wb.Worksheets
        If ws.Name <> tocName And ws.Visible = xlSheetVisible Then
            Set cellA = toc.Cells(rowIdx, 1)
            cellA.Value = ws.Name
            cellA.Font.Color = RGB(5, 99, 193)
            cellA.Font.Underline = xlUnderlineStyleSingle
            toc.Hyperlinks.Add Anchor:=cellA, Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            On Error Resume Next
            desc = ws.Range(descCell).Value
            On Error GoTo ErrHandler
            Set cellB = toc.Cells(rowIdx, 2)
            If Not IsEmpty(desc) Then cellB.Value = CStr(desc)
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
    MsgBox "DM_GenerateDeriveTOC Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
"""

# ================================================================
# Module 5: modSheetSplit
# ================================================================
MOD5 = """Attribute VB_Name = "modSheetSplit"
Option Explicit

'============================================================================
' 模块: modSheetSplit
' 描述: CDM工作表拆分 — 按列值将一个sheet拆分为多个sheet
'       Sheet Splitter — split one sheet into many by column value
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无
'============================================================================

Private Function DM_MakeSafeSheetName(ByVal name As String) As String
    Dim result As String, illegalChars As String, i As Long
    illegalChars = "\\/*?:[]"
    result = name
    For i = 1 To Len(illegalChars)
        result = Replace(result, Mid(illegalChars, i, 1), "_")
    Next i
    If Len(result) > 31 Then result = Left(result, 31)
    If Len(result) = 0 Then result = "Sheet"
    DM_MakeSafeSheetName = result
End Function

Private Function DM_SheetExistsLocal(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    DM_SheetExistsLocal = (Err.Number = 0)
    Err.Clear
End Function

Public Function DM_SplitSheetByColumn(ByVal ws As Worksheet, _
    Optional ByVal splitCol As String = "A", _
    Optional ByVal headerRows As Long = 1) As String()
    On Error GoTo ErrHandler
    Dim wb As Workbook, lastRow As Long, lastCol As Long, splitColIdx As Long
    Dim i As Long, key As Variant, groups As Object, rowColl As Collection
    Dim child As Worksheet, rowIdx As Long, colIdx As Long
    Dim srcRow As Long, destRow As Long, result() As String, k As Long
    Dim safeName As String, conflictNames As String, v As Variant

    Set wb = ws.Parent
    splitColIdx = ws.Range(splitCol & "1").Column
    lastRow = ws.Cells(ws.Rows.Count, splitColIdx).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow <= headerRows Then
        ReDim result(0 To -1): DM_SplitSheetByColumn = result: Exit Function
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    For i = headerRows + 1 To lastRow
        key = CStr(ws.Cells(i, splitColIdx).Value)
        If Len(key) = 0 Then key = "(blank)"
        If Not groups.Exists(key) Then
            Set rowColl = New Collection: groups.Add rowColl, key
        End If
        groups(key).Add i
    Next i

    For Each key In groups.Keys
        safeName = DM_MakeSafeSheetName(CStr(key))
        If DM_SheetExistsLocal(wb, safeName) Then
            conflictNames = conflictNames & safeName & vbCrLf
        End If
    Next key
    If Len(conflictNames) > 0 Then
        MsgBox "Sheet names already exist:" & vbCrLf & conflictNames, _
            vbExclamation, "CDM Toolkit - Split Conflict"
        ReDim result(0 To -1): DM_SplitSheetByColumn = result: Exit Function
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ReDim result(0 To groups.Count - 1): k = 0

    For Each key In groups.Keys
        safeName = DM_MakeSafeSheetName(CStr(key))
        Set child = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        child.Name = safeName: result(k) = safeName: k = k + 1
        destRow = 1
        If headerRows > 0 Then
            For srcRow = 1 To headerRows
                For colIdx = 1 To lastCol
                    child.Cells(destRow, colIdx).Value = ws.Cells(srcRow, colIdx).Value
                    child.Cells(destRow, colIdx).Font.Bold = True
                Next colIdx
                destRow = destRow + 1
            Next srcRow
        End If
        Set rowColl = groups(key)
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
    MsgBox "DM_SplitSheetByColumn Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Function DM_SplitSheetWithProgress(ByVal ws As Worksheet, _
    Optional ByVal splitCol As String = "A", _
    Optional ByVal headerRows As Long = 1) As String()
    Dim oldStatusBar As String
    oldStatusBar = Application.StatusBar
    Application.StatusBar = "CDM Toolkit: Splitting sheet..."
    DM_SplitSheetWithProgress = DM_SplitSheetByColumn(ws, splitCol, headerRows)
    Application.StatusBar = oldStatusBar
End Function
"""

# ================================================================
# Write all modules
# ================================================================
modules = {
    "modTextStandardise.bas": MOD1,
    "modHighlightDup.bas": MOD2,
    "modBlankRows.bas": MOD3,
    "modTOC.bas": MOD4,
    "modSheetSplit.bas": MOD5,
}

for name, content in modules.items():
    path = os.path.join(MOD_DIR, name)
    with open(path, 'w', encoding='gbk') as f:
        f.write(content)
    print(f"Written: {name}")

print("\nFirst 5 modules done! Run part2 for remaining modules.")
