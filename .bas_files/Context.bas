Attribute VB_Name = "Context"
Option Explicit

Public Function LoadCfgTemplates(ByVal wb As Workbook) As Object
    Dim ws As Worksheet
    Dim cfg As Object
    Dim lastRow As Long
    Dim rowNo As Long
    Dim code As String
    Dim rowDict As Object

    Set ws = wb.Sheets("TEMPLATES")
    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = 1

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For rowNo = 2 To lastRow
        code = Trim$(CellText(ws.Cells(rowNo, "B").Value))
        If Len(code) > 0 Then
            Set rowDict = CreateObject("Scripting.Dictionary")
            rowDict.CompareMode = 1
            rowDict("selected") = ParseEnabled(ws.Cells(rowNo, "A").Value)
            rowDict("template_code") = code
            rowDict("description") = CellText(ws.Cells(rowNo, "C").Value)
            rowDict("docx_file") = CellText(ws.Cells(rowNo, "D").Value)
            rowDict("file_prefix") = CellText(ws.Cells(rowNo, "E").Value)
            cfg(code) = rowDict
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

    Set ws = wb.Sheets("INPUT")
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.CompareMode = 1

    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)

    For rowNo = 1 To lastRow
        keyText = Trim$(CellText(ws.Cells(rowNo, "A").Value))
        If Len(keyText) > 0 Then
            valueText = CellText(ws.Cells(rowNo, "D").Value)
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
