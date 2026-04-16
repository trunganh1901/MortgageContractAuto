Attribute VB_Name = "Context"
Option Explicit

' Purpose: Read the TEMPLATES sheet into a dictionary of template configuration dictionaries.
'          Reads the entire data range as a 2-D array to avoid per-cell COM round-trips.
' Inputs:  wb = current workbook.
' Outputs: Dictionary keyed by template_code; each value is a sub-dictionary with
'          selected, template_code, description, docx_file, file_prefix.
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

    If lastRow < 2 Then
        Set LoadCfgTemplates = cfg
        Exit Function
    End If

    ' Read columns A-E in one array read (much faster than per-cell access)
    data = ws.Range("A2:E" & lastRow).Value

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

' Purpose: Read the INPUT sheet into a context dictionary (column A = key, column D = value).
'          Reads both columns as arrays to avoid per-cell COM round-trips.
' Inputs:  wb = current workbook.
' Outputs: Dictionary keyed by column-A values; values are the formatted column-D text.
Public Function BuildContext(ByVal wb As Workbook) As Object
    Dim ws As Worksheet
    Dim ctx As Object
    Dim lastRow As Long
    Dim rowNo As Long
    Dim keyText As String
    Dim valueText As String
    Dim dataA As Variant
    Dim dataD As Variant

    Set ws = wb.Sheets("INPUT")
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.CompareMode = 1

    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)

    If lastRow < 1 Then
        Set BuildContext = ctx
        Exit Function
    End If

    ' Read both columns as arrays in two round-trips instead of 2*lastRow round-trips
    dataA = ws.Range("A1:A" & lastRow).Value
    dataD = ws.Range("D1:D" & lastRow).Value

    For rowNo = 1 To lastRow
        keyText = Trim$(CellText(dataA(rowNo, 1)))
        If Len(keyText) > 0 Then
            valueText = ExcelCellText(dataD(rowNo, 1))
            If Len(valueText) > 0 Then
                ctx(keyText) = valueText
            End If
        End If
    Next rowNo

    Set BuildContext = ctx
End Function

' Purpose: Interpret a cell value as a boolean enabled/disabled flag.
' Inputs:  value = raw cell value (Boolean, numeric, or text).
' Outputs: True when the value represents an enabled/selected state.
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
