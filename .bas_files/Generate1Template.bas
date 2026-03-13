Attribute VB_Name = "Generate1Template"
Option Explicit
Sub GenerateTemplateFromSheet()

    Dim wsUI As Worksheet
    Dim oldTemplate As String
    Dim oldSource As String
    Dim templateCode As String
    Dim sheetName As String

    sheetName = ActiveSheet.Name

    If Left(sheetName, 4) <> "TPL_" Then
        MsgBox "This button must be used on a TPL_* sheet.", vbExclamation
        Exit Sub
    End If

    templateCode = Mid(sheetName, 5)   ' remove "TPL_"

    Set wsUI = ThisWorkbook.Sheets("UI_DASHBOARD")
    
    'Disable buttons on this TPL sheet
    DisableAllButtons ActiveSheet
    Application.Cursor = xlWait
    
    On Error GoTo CLEANUP
    Err.Clear
    
    ' Save state
    oldTemplate = wsUI.Range("B2").Value
    oldSource = wsUI.Range("B8").Value

    ' Force template + source
    wsUI.Range("B2").Value = templateCode
    wsUI.Range("B8").Value = sheetName

    ' Run generator
    GeneralDocGenerate
    
CLEANUP:
    ' Restore state
    wsUI.Range("B2").Value = oldTemplate
    wsUI.Range("B8").Value = oldSource
    
    ' ?? Re-enable buttons
    EnableAllButtons ActiveSheet
    Application.Cursor = xlDefault

    If Err.Number <> 0 Then
        MsgBox "Generation failed.", vbCritical
    End If
    

End Sub
