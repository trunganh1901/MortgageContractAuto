Attribute VB_Name = "BBGN_PXK_from_others"
Sub GenerateBBGN_PXK_FromActiveSheet()

    Dim wsUI As Worksheet
    Dim oldTemplate As String
    Dim oldSource As String

    Set wsUI = ThisWorkbook.Sheets("UI_DASHBOARD")

    ' Save previous state
    oldTemplate = wsUI.Range("B2").Value
    oldSource = wsUI.Range("B8").Value

    ' Force BBGN_PXK template + active sheet as data source
    wsUI.Range("B2").Value = "BBGN_PXK"
    wsUI.Range("B8").Value = ActiveSheet.Name

    ' Run generator
    GeneralDocGenerate
        
    'Success popup
    MsgBox "Contracts generated successfully!" & vbCrLf, vbInformation, "Done"
        
    ' Restore previous state
    wsUI.Range("B2").Value = oldTemplate
    wsUI.Range("B8").Value = oldSource

End Sub
