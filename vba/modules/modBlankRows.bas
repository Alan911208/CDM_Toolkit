Attribute VB_Name = "modBlankRows"
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
