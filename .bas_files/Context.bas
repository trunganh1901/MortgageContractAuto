Attribute VB_Name = "Context"
Option Explicit

Public Function LoadCfgTemplates(ByVal wb As Workbook) As Object
    Dim ws As Worksheet
    Dim cfg As Object
    Dim lastRow As Long
    Dim rowNo As Long
    Dim code As String
    Dim rowDict As Object
    Dim data As Variant

    Set ws = wb.Sheets("TEMPLATES")
    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = 1

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then Set LoadCfgTemplates = cfg: Exit Function

    data = ws.Range(ws.Cells(2, "A"), ws.Cells(lastRow, "E")).Value

    For rowNo = 1 To lastRow - 1
        code = Trim$(CellText(data(rowNo, 2)))
        If Len(code) > 0 Then
            Set rowDict = CreateObject("Scripting.Dictionary")
            rowDict.CompareMode = 1
            rowDict("selected") = ParseEnabled(data(rowNo, 1))
            rowDict("template_code") = code
            rowDict("description") = CellText(data(rowNo, 3))
            rowDict("docx_file") = CellText(data(rowNo, 4))
            rowDict("file_prefix") = CellText(data(rowNo, 5))
            Set cfg(code) = rowDict
        End If
    Next rowNo

    Set LoadCfgTemplates = cfg
End Function

Public Function BuildContext(ByVal wb As Workbook) As Object
    Dim ws As Worksheet
    Dim ctx As Object
    Dim lastRow As Long
    Dim rowNo As Long
    Dim keyText As String
    Dim valueText As String
    Dim keysData As Variant
    Dim valuesData As Variant

    Set ws = wb.Sheets("INPUT")
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.CompareMode = 1

    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)

    If lastRow < 1 Then Set BuildContext = ctx: Exit Function

    keysData = ws.Range(ws.Cells(1, "A"), ws.Cells(lastRow, "A")).Value
    valuesData = ws.Range(ws.Cells(1, "D"), ws.Cells(lastRow, "D")).Value

    For rowNo = 1 To lastRow
        keyText = Trim$(CellText(keysData(rowNo, 1)))
        If Len(keyText) > 0 Then
            valueText = ExcelCellText(valuesData(rowNo, 1))
            If Len(valueText) > 0 Then
                ctx(keyText) = valueText
            End If
        End If
    Next rowNo

    Set BuildContext = ctx
End Function

Private Function ParseEnabled(ByVal value As Variant) As Boolean
    Dim textValue As String

    If VarType(value) = vbBoolean Then
        ParseEnabled = CBool(value)
        Exit Function
    End If

    textValue = UCase$(Trim$(CellText(value)))
    Select Case textValue
        Case "1", "TRUE", "YES", "Y", "X"
            ParseEnabled = True
    End Select
End Function
