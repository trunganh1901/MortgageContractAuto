Attribute VB_Name = "modExportDocument"
Option Explicit

Public Function ExportDocument() As String
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim cfg As Object
    Dim ctx As Object
    Dim templateCode As Variant
    Dim templateCfg As Object
    Dim lastOutputPath As String

    Set wb = ThisWorkbook
    Set cfg = LoadCfgTemplates(wb)
    Set ctx = BuildContext(wb)

    For Each templateCode In cfg.Keys
        Set templateCfg = cfg(CStr(templateCode))
        If GetDictBoolean(templateCfg, "selected") Then
            lastOutputPath = RenderTemplate(templateCfg, ctx, wb)
        End If
    Next templateCode

    If Len(lastOutputPath) = 0 Then
        Err.Raise vbObjectError + 514, "ExportDocument", "No templates selected on the TEMPLATES sheet."
    End If

    ExportDocument = lastOutputPath
    PromptOpenOutputFolder lastOutputPath
    Exit Function

ErrorHandler:
    MsgBox "Export failed: " & Err.Description, vbCritical, "ExportDocument"
End Function

Public Function RunContractAutomation() As String
    RunContractAutomation = ExportDocument()
End Function
