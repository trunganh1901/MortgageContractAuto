Attribute VB_Name = "Context"
Option Explicit

Public Function LoadCfgTemplates(ByVal wb As Workbook) As Object
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim cfg As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowNo As Long
    Dim colNo As Long
    Dim code As String
    Dim rowDict As Object

    Set ws = wb.Sheets("CFG_TEMPLATES")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = 1

    For colNo = 1 To lastCol
        headerMap(Trim$(CellText(ws.Cells(1, colNo).Value))) = colNo
    Next colNo

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = 1

    For rowNo = 2 To lastRow
        code = Trim$(GetCellByHeader(ws, rowNo, headerMap, "TemplateCode"))
        If Len(code) > 0 Then
            Set rowDict = CreateObject("Scripting.Dictionary")
            rowDict.CompareMode = 1
            rowDict("excel_sheet") = GetCellByHeader(ws, rowNo, headerMap, "ExcelSheet")
            rowDict("docx_file") = GetCellByHeader(ws, rowNo, headerMap, "DocxFile")
            rowDict("file_prefix") = GetCellByHeader(ws, rowNo, headerMap, "FilePrefix")
            rowDict("description") = GetCellByHeader(ws, rowNo, headerMap, "Description")
            rowDict("enabled") = ParseEnabled(GetCellByHeader(ws, rowNo, headerMap, "Enabled"))
            cfg(code) = rowDict
        End If
    Next rowNo

    Set LoadCfgTemplates = cfg
End Function

Public Function LoadItems(ByVal wb As Workbook) As Collection
    Dim items As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headers() As String
    Dim rowNo As Long
    Dim colNo As Long
    Dim item As Object
    Dim soLuong As Double
    Dim donGia As Double
    Dim thanhTien As Double
    Dim rowHasData As Boolean

    Set items = New Collection

    On Error Resume Next
    Set ws = wb.Sheets("Items")
    On Error GoTo 0
    If ws Is Nothing Then
        Set LoadItems = items
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol = 0 Then
        Set LoadItems = items
        Exit Function
    End If

    ReDim headers(1 To lastCol)
    For colNo = 1 To lastCol
        headers(colNo) = NormalizeHeader(CellText(ws.Cells(1, colNo).Value))
    Next colNo

    For rowNo = 2 To lastRow
        rowHasData = False
        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = 1

        For colNo = 1 To lastCol
            item(headers(colNo)) = CellText(ws.Cells(rowNo, colNo).Value)
            If Len(Trim$(item(headers(colNo)))) > 0 Then rowHasData = True
        Next colNo

        If rowHasData Then
            soLuong = ToNumber(GetDictString(item, "so_luong"))
            donGia = ToNumber(GetDictString(item, "don_gia"))
            thanhTien = ToNumber(GetDictString(item, "thanh_tien"))
            If thanhTien = 0# Then thanhTien = soLuong * donGia

            item("so_luong_num") = soLuong
            item("don_gia_num") = donGia
            item("thanh_tien_num") = thanhTien
            item("so_luong") = FormatVN(soLuong, True)
            item("don_gia") = FormatVN(donGia)
            item("thanh_tien") = FormatVN(thanhTien)

            items.Add item
        End If
    Next rowNo

    Set LoadItems = items
End Function

Public Function BuildContext(ByVal wb As Workbook, ByVal sheetName As String) As Object
    Dim ws As Worksheet
    Dim ctx As Object
    Dim lastRow As Long
    Dim rowNo As Long
    Dim keyText As String
    Dim valueText As String

    Set ws = wb.Sheets(sheetName)
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.CompareMode = 1

    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, 4).End(xlUp).Row)

    For rowNo = 1 To lastRow
        keyText = Trim$(CellText(ws.Cells(rowNo, 1).Value))
        If Len(keyText) > 0 Then
            valueText = CellText(ws.Cells(rowNo, 4).Value)
            ctx(keyText) = valueText
        End If
    Next rowNo

    If Not ctx.Exists("DAY") Then ctx("DAY") = Day(Date)
    If Not ctx.Exists("MONTH") Then ctx("MONTH") = Month(Date)
    If Not ctx.Exists("YEAR") Then ctx("YEAR") = Year(Date)

    Set BuildContext = ctx
End Function

Public Sub EnrichTotals(ByVal ctx As Object, ByVal items As Collection)
    Dim grand As Double
    Dim rate As Double
    Dim vatAmount As Double
    Dim totalWithVat As Double
    Dim item As Variant

    grand = 0#
    For Each item In items
        grand = grand + CDbl(item("thanh_tien_num"))
    Next item

    rate = ToNumber(GetDictString(ctx, "VAT_RATE", CStr(DEFAULT_VAT_RATE)))
    If rate = 0# Then rate = DEFAULT_VAT_RATE

    CalculateVAT grand, rate, vatAmount, totalWithVat

    ctx("items") = items
    ctx("grand_total") = RoundHalfUpValue(grand, 0)
    ctx("grand_total_formatted") = FormatVN(grand)
    ctx("vat_amount_formatted") = FormatVN(vatAmount)
    ctx("grand_total_vat_formatted") = FormatVN(totalWithVat)
    ctx("grand_total_text") = VndToWords(CLng(RoundHalfUpValue(grand, 0)))
    ctx("grand_total_vat_text") = VndToWords(CLng(RoundHalfUpValue(totalWithVat, 0)))
End Sub

Public Sub CalculateVAT(ByVal grandTotal As Double, ByVal vatRate As Double, ByRef vatAmount As Double, ByRef totalWithVat As Double)
    vatAmount = RoundHalfUpValue(grandTotal * vatRate, 0)
    totalWithVat = RoundHalfUpValue(grandTotal + vatAmount, 0)
End Sub

Private Function GetCellByHeader(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerMap As Object, ByVal headerName As String) As String
    If headerMap.Exists(headerName) Then
        GetCellByHeader = CellText(ws.Cells(rowNo, CLng(headerMap(headerName))).Value)
    Else
        GetCellByHeader = vbNullString
    End If
End Function

Private Function NormalizeHeader(ByVal value As String) As String
    Dim textValue As String
    textValue = LCase$(Trim$(value))
    textValue = Replace$(textValue, " ", "_")
    NormalizeHeader = textValue
End Function

Private Function ParseEnabled(ByVal value As String) As Boolean
    Select Case UCase$(Trim$(value))
        Case "1", "TRUE", "YES", "Y"
            ParseEnabled = True
    End Select
End Function
