Attribute VB_Name = "modHighlightDup"
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
