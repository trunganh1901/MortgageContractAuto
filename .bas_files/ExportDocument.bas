Attribute VB_Name = "modExportDocument"
Option Explicit

Public Function ExportDocument() As String
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim cfg As Object
    Dim ctx As Object
    Dim logEntry As Object
    Dim exportOutputs As Collection
    Dim templateCode As Variant
    Dim templateCfg As Object
    Dim lastOutputPath As String
    Dim startedAt As Date

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Exporting document..."

    Set wb = ThisWorkbook
    Set cfg = LoadCfgTemplates(wb)
    Set ctx = BuildContext(wb)
    startedAt = Now
    Set logEntry = CreateExportLog(wb, ctx, startedAt)
    Set exportOutputs = New Collection

    For Each templateCode In cfg.Keys
        Set templateCfg = cfg(CStr(templateCode))
        If GetDictBoolean(templateCfg, "selected") Then
            Application.StatusBar = "Rendering: " & GetDictString(templateCfg, "description")
            lastOutputPath = RenderTemplate(templateCfg, ctx, wb)
            exportOutputs.Add CreateExportOutput(templateCfg, lastOutputPath)
        End If
    Next templateCode

    If Len(lastOutputPath) = 0 Then
        Err.Raise vbObjectError + 514, "ExportDocument", "No templates selected on the TEMPLATES sheet."
    End If

    SaveExportLog logEntry, exportOutputs, Now, "success"

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ExportDocument = lastOutputPath
    PromptOpenOutputFolder lastOutputPath
    Exit Function

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not logEntry Is Nothing Then
        If exportOutputs Is Nothing Then Set exportOutputs = New Collection
        SaveExportLog logEntry, exportOutputs, Now, "failed", Err.Description
    End If
    MsgBox "Export failed: " & Err.Description, vbCritical, "ExportDocument"
End Function

Public Function RunContractAutomation() As String
    RunContractAutomation = ExportDocument()
End Function
