Attribute VB_Name = "GenerateALL"
Option Explicit
Sub GenerateAllTemplates()

    Dim wsUI As Worksheet
    Dim oldTemplate As String
    Dim oldSource As String
    
    Set wsUI = ThisWorkbook.Sheets("UI_DASHBOARD")

    'Disable buttons on this TPL sheet
    DisableAllButtons ActiveSheet
    Application.Cursor = xlWait
    
    On Error GoTo CLEANUP

    ' Save previous state
    oldTemplate = wsUI.Range("B2").Value
    oldSource = wsUI.Range("B8").Value

    ' Force ALL + current sheet as source
    wsUI.Range("B2").Value = "ALL"
    wsUI.Range("B8").Value = ActiveSheet.Name

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
