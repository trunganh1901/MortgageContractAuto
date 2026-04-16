Attribute VB_Name = "modExportDocument"
Option Explicit

' Purpose: Main export entry point. Reads config and context, renders all selected templates
'          using one shared Word instance, logs the run, and prompts the user to open the output folder.
' Inputs:  None (reads TEMPLATES and INPUT sheets from ThisWorkbook).
' Outputs: Full path of the last rendered DOCX file, or empty string on failure.
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
    Dim wordApp As Object
    Dim createdWord As Boolean
    Dim selectedCount As Long
    Dim processedCount As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook

    ShowStatus "Loading templates and input context..."
    Set cfg = LoadCfgTemplates(wb)
    Set ctx = BuildContext(wb)
    startedAt = Now
    Set logEntry = CreateExportLog(wb, ctx, startedAt)
    Set exportOutputs = New Collection

    ' Count selected templates before starting Word
    For Each templateCode In cfg.Keys
        If GetDictBoolean(cfg(CStr(templateCode)), "selected") Then
            selectedCount = selectedCount + 1
        End If
    Next templateCode

    If selectedCount = 0 Then
        Err.Raise vbObjectError + 514, "ExportDocument", "No templates selected on the TEMPLATES sheet."
    End If

    ' Create one Word instance shared across all template renders
    ShowStatus "Starting Microsoft Word..."
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo ErrorHandler
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
        createdWord = True
    End If
    wordApp.Visible = False

    For Each templateCode In cfg.Keys
        Set templateCfg = cfg(CStr(templateCode))
        If GetDictBoolean(templateCfg, "selected") Then
            processedCount = processedCount + 1
            ShowStatus "Rendering " & processedCount & " / " & selectedCount & _
                       ": " & GetDictString(templateCfg, "description") & "..."
            lastOutputPath = RenderTemplate(templateCfg, ctx, wb, wordApp)
            exportOutputs.Add CreateExportOutput(templateCfg, lastOutputPath)
        End If
    Next templateCode

    ' Quit Word only if we started it
    If createdWord Then
        On Error Resume Next
        wordApp.Quit wdDoNotSaveChanges
        On Error GoTo ErrorHandler
    End If
    Set wordApp = Nothing

    ShowStatus "Saving audit log..."
    SaveExportLog logEntry, exportOutputs, Now, "success"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ClearStatus

    ExportDocument = lastOutputPath
    PromptOpenOutputFolder lastOutputPath
    Exit Function

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ClearStatus
    If Not wordApp Is Nothing Then
        On Error Resume Next
        If createdWord Then wordApp.Quit wdDoNotSaveChanges
        On Error GoTo 0
        Set wordApp = Nothing
    End If
    If Not logEntry Is Nothing Then
        If exportOutputs Is Nothing Then Set exportOutputs = New Collection
        SaveExportLog logEntry, exportOutputs, Now, "failed", Err.Description
    End If
    MsgBox "Export failed: " & Err.Description, vbCritical, "ExportDocument"
End Function

Public Function RunContractAutomation() As String
    RunContractAutomation = ExportDocument()
End Function
