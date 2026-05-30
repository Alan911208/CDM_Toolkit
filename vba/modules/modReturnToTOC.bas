Attribute VB_Name = "modReturnToTOC"
Option Explicit

' Module: modReturnToTOC
' Purpose: One-click return to TOC navigation
'   - DM_GoToTOC: Jump to TOC sheet (can assign Ctrl+Shift+T shortcut)
'   - DM_AddReturnToTOCButton: Add "Back to TOC" button to current sheet
'   - DM_AddReturnToTOCButtonAll: Add button to all sheets at once
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Private Const BTN_NAME As String = "CDM_BackToTOC"

Public Sub DM_GoToTOC(Optional ByVal tocName As String = "TOC")
    On Error GoTo ErrHandler

    Dim wb As Workbook, toc As Worksheet

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set toc = wb.Worksheets(tocName)
    If Err.Number <> 0 Then
        Err.Clear
        Set toc = wb.Worksheets("TOC")
        If Err.Number <> 0 Then
            Err.Clear
            Set toc = wb.Worksheets("Table of Contents")
            If Err.Number <> 0 Then
                MsgBox "TOC sheet not found. Generate TOC first.", vbInformation, "CDM Toolkit"
                Exit Sub
            End If
        End If
    End If
    On Error GoTo ErrHandler

    toc.Activate
    toc.Range("A1").Select
    Exit Sub

ErrHandler:
    MsgBox "DM_GoToTOC Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_AddReturnToTOCButton(Optional ByVal ws As Worksheet = Nothing, _
    Optional ByVal tocName As String = "TOC")

    On Error GoTo ErrHandler

    Dim shp As Shape

    If ws Is Nothing Then Set ws = ActiveSheet
    If ws.Name = tocName Then Exit Sub

    On Error Resume Next
    ws.Shapes(BTN_NAME).Delete
    On Error GoTo ErrHandler

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 5, 2, 80, 22)
    shp.Name = BTN_NAME
    shp.Fill.ForeColor.RGB = RGB(68, 114, 196)
    shp.Fill.Solid
    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Text = "Back to TOC"
    shp.TextFrame2.TextRange.Font.Size = 9
    shp.TextFrame2.TextRange.Font.Bold = True
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

    shp.OnAction = "DM_GoToTOC_Click"

    Exit Sub

ErrHandler:
    MsgBox "DM_AddReturnToTOCButton Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_GoToTOC_Click()
    DM_GoToTOC
End Sub

Public Sub DM_AddReturnToTOCButtonAll(Optional ByVal wb As Workbook = Nothing, _
    Optional ByVal tocName As String = "TOC")

    On Error GoTo ErrHandler
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook

    Application.ScreenUpdating = False

    For Each ws In wb.Worksheets
        If ws.Name <> tocName And ws.Visible = xlSheetVisible Then
            DM_AddReturnToTOCButton ws, tocName
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox "Return buttons added to all sheets.", vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_AddReturnToTOCButtonAll Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
