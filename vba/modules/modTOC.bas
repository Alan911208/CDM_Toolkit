Attribute VB_Name = "modTOC"
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
