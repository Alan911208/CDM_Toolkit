"""Generate 3 new VBA modules with proper ASCII encoding."""
import os

MOD_DIR = r"e:\Project\Tools\Tools_DM\vba\modules"
os.makedirs(MOD_DIR, exist_ok=True)

def write_bas(filename, content):
    path = os.path.join(MOD_DIR, filename)
    with open(path, 'w', encoding='ascii') as f:
        f.write(content)
    print(f"Written: {filename}")

# ================================================================
# 1. modSheetCreate.bas — Batch create sheets from name list
# ================================================================
write_bas("modSheetCreate.bas", """\
Attribute VB_Name = "modSheetCreate"
Option Explicit

' Module: modSheetCreate
' Purpose: Batch create empty sheets from a name list
'   - DM_CreateSheetsFromList: Create multiple sheets from comma/line/array input
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Public Function DM_CreateSheetsFromList(ByVal wb As Workbook, _
    ByVal sheetNames As Variant, _
    Optional ByVal clearFirst As Boolean = False) As Long

    On Error GoTo ErrHandler

    Dim names() As String, i As Long, nameStr As String
    Dim created As Long, ws As Worksheet

    If IsArray(sheetNames) Then
        ReDim names(0 To UBound(sheetNames))
        For i = 0 To UBound(sheetNames)
            names(i) = CStr(sheetNames(i))
        Next i
    Else
        nameStr = CStr(sheetNames)
        nameStr = Replace(nameStr, vbCrLf, ",")
        nameStr = Replace(nameStr, vbCr, ",")
        nameStr = Replace(nameStr, vbLf, ",")
        names = Split(nameStr, ",")
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 0 To UBound(names)
        nameStr = Trim(names(i))
        If Len(nameStr) = 0 Then GoTo NextName

        If Len(nameStr) > 31 Then nameStr = Left(nameStr, 31)

        Dim illegal As String, j As Long
        illegal = "\\/*?:[]"
        For j = 1 To Len(illegal)
            nameStr = Replace(nameStr, Mid(illegal, j, 1), "_")
        Next j

        If Len(nameStr) = 0 Then GoTo NextName

        On Error Resume Next
        Set ws = wb.Worksheets(nameStr)
        If Err.Number = 0 Then
            If clearFirst Then
                ws.Delete
            Else
                Dim suffix As Long: suffix = 1
                Dim base As String: base = nameStr
                Do While Err.Number = 0
                    Err.Clear
                    nameStr = Left(base, 28) & "_" & suffix
                    suffix = suffix + 1
                    Set ws = wb.Worksheets(nameStr)
                Loop
            End If
        End If
        Err.Clear
        On Error GoTo ErrHandler

        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nameStr
        created = created + 1

NextName:
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    DM_CreateSheetsFromList = created
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "DM_CreateSheetsFromList Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function
""")

# ================================================================
# 2. modReturnToTOC.bas — One-click return to TOC
# ================================================================
write_bas("modReturnToTOC.bas", """\
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
""")

print("Module 1-2 done!")
