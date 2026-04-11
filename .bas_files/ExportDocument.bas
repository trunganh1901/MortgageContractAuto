Attribute VB_Name = "modExportDocument"
Option Explicit

Public Function RunContractAutomation() As String
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim cfg As Object
    Dim templateCode As String
    Dim folder As String

    Set wb = ThisWorkbook
    Set cfg = LoadCfgTemplates(wb)
    templateCode = Trim$(CellText(wb.Sheets("UI_DASHBOARD").Range("B2").Value))

    folder = RunContractWorkflow(wb, templateCode, cfg)
    wb.Sheets("UI_DASHBOARD").Range("B7").Value = folder
    RunContractAutomation = folder
    Exit Function

ErrorHandler:
    MsgBox "Export failed: " & Err.Description, vbCritical, "ExportDocument"
End Function

Public Function ExportDocument() As String
    ExportDocument = RunContractAutomation()
End Function

Public Function RunContractWorkflow(ByVal wb As Workbook, ByVal templateCode As String, ByVal cfg As Object) As String
    On Error GoTo ErrorHandler

    Dim overrideSheet As String
    Dim items As Collection
    Dim keys As Variant
    Dim keyValue As Variant
    Dim templateCfg As Object
    Dim ctx As Object
    Dim seq As String
    Dim customerName As String
    Dim lastFolder As String
    Dim sourceSheet As String

    overrideSheet = Trim$(CellText(wb.Sheets("UI_DASHBOARD").Range("B8").Value))
    Set items = LoadItems(wb)

    If UCase$(templateCode) = "ALL" Then
        keys = cfg.Keys
    Else
        ReDim keys(0 To 0)
        keys(0) = templateCode
    End If

    lastFolder = vbNullString

    For Each keyValue In keys
        If cfg.Exists(CStr(keyValue)) Then
            Set templateCfg = cfg(CStr(keyValue))
            If GetDictBoolean(templateCfg, "enabled") Then
                sourceSheet = overrideSheet
                If Len(sourceSheet) = 0 Then
                    sourceSheet = CellText(templateCfg("excel_sheet"))
                End If

                Set ctx = BuildContext(wb, sourceSheet)
                seq = NormalizeSequence(GetDictString(ctx, "STT_HD", "00"))
                EnrichTotals ctx, items

                customerName = GetDictString(ctx, "TEN_KH")
                If Len(customerName) = 0 Then customerName = GetDictString(ctx, "KH_ABB")
                If Len(customerName) = 0 Then customerName = "contract"

                lastFolder = RenderTemplate(templateCfg, ctx, seq, customerName, wb)
            End If
        End If
    Next keyValue

    RunContractWorkflow = lastFolder
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "RunContractWorkflow", Err.Description
End Function
