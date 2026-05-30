Attribute VB_Name = "CDM_Toolkit"
Option Explicit

' Module: CDM_Toolkit
' Purpose: CDM Toolkit Master Module - version info and quick-launch entry points
'
'   Sub-modules (16 total):
'     [Data Cleaning]   modTextStandardise, modHighlightDup, modBlankRows, modStrikethrough
'     [Document Mgmt]   modTOC, modSheetSplit, modWorkbookMerge, modReturnToTOC
'     [Sheet Mgmt]      modSheetManager, modSheetCreate
'     [Audit & Review]  modTrackChanges, modSpecialChars
'     [Calculation]     modDateCalc, modUnitConvert, modCreatinineClearance
'     [Utilities]       modFileListing
'
'   Quick Launch:
'     DM_QuickClean  - Text standardise + highlight dups + delete blank rows
'     DM_QuickDoc    - Generate TOC + add return buttons to all sheets
'     DM_About       - Version info
'
' Version: 1.1 / 2026-05-30
' Install: Alt+F11 -> File -> Import File -> select all vba/modules/*.bas

Public Const CDM_TOOLKIT_VERSION As String = "1.1"
Public Const CDM_TOOLKIT_DATE As String = "2026-05-30"

Public Sub DM_About()
    MsgBox "CDM Excel Toolkit v" & CDM_TOOLKIT_VERSION & vbCrLf & vbCrLf & _
        "Clinical Data Management Excel Macro Toolkit" & vbCrLf & vbCrLf & _
        "16 modules covering:" & vbCrLf & _
        "  [Data Cleaning]   Text standardise, Highlight dups, Blank rows, Strikethrough" & vbCrLf & _
        "  [Document Mgmt]   TOC, Sheet split, Workbook merge, Return to TOC" & vbCrLf & _
        "  [Sheet Mgmt]      Sheet rename/list/create/delete" & vbCrLf & _
        "  [Audit & Review]  Track changes, Special char scanner" & vbCrLf & _
        "  [Calculation]     Date calc, Unit convert, Creatinine clearance" & vbCrLf & _
        "  [Utilities]       File listing" & vbCrLf & vbCrLf & _
        "Date: " & CDM_TOOLKIT_DATE, _
        vbInformation, "CDM Toolkit"
End Sub

Public Sub DM_QuickClean(ByVal ws As Worksheet)
    On Error GoTo ErrHandler

    Dim count1 As Long, count3 As Long, count4 As Long

    Application.StatusBar = "CDM Toolkit: Standardising terms..."
    count1 = DM_TextStandardise(ws)

    Application.StatusBar = "CDM Toolkit: Highlighting duplicates..."
    DM_HighlightDuplicates ws

    Application.StatusBar = "CDM Toolkit: Deleting blank rows..."
    count3 = DM_DeleteBlankRows(ws)

    Application.StatusBar = "CDM Toolkit: Deleting strikethrough rows..."
    count4 = DM_DeleteStrikethroughRows(ws, False)

    Application.StatusBar = False
    MsgBox "Quick Clean complete!" & vbCrLf & _
        "  Terms replaced: " & count1 & vbCrLf & _
        "  Blank rows removed: " & count3 & vbCrLf & _
        "  Strikethrough rows removed: " & count4, _
        vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "DM_QuickClean Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_QuickDoc(ByVal wb As Workbook)
    On Error GoTo ErrHandler

    Application.StatusBar = "CDM Toolkit: Generating TOC..."
    DM_GenerateTOC wb

    Application.StatusBar = "CDM Toolkit: Adding return buttons..."
    DM_AddReturnToTOCButtonAll wb

    Application.StatusBar = False
    MsgBox "Document setup complete! TOC generated with return buttons on all sheets.", _
        vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "DM_QuickDoc Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
