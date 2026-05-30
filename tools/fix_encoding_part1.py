"""
Fix encoding: regenerate all VBA .bas files with proper UTF-8 encoding.
"""
import os

VBA_DIR = r"e:\Project\Tools\Tools_DM\vba"
MOD_DIR = os.path.join(VBA_DIR, "modules")
os.makedirs(MOD_DIR, exist_ok=True)

def write_bas(filename, content):
    path = os.path.join(MOD_DIR, filename)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"  OK: {filename}")

# ============================================================
# modTextStandardise.bas
# ============================================================
write_bas("modTextStandardise.bas", '''Attribute VB_Name = "modTextStandardise"
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
''')

# ============================================================
# modHighlightDup.bas
# ============================================================
write_bas("modHighlightDup.bas", '''Attribute VB_Name = "modHighlightDup"
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
''')

# ============================================================
# modBlankRows.bas
# ============================================================
write_bas("modBlankRows.bas", '''Attribute VB_Name = "modBlankRows"
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
''')

print("First 3 modules done.")
