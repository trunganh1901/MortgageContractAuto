Attribute VB_Name = "PythonPort"
Option Explicit

Private Const DEFAULT_VAT_RATE As Double = 0.08
Private Const wdFindContinue As Long = 1
Private Const wdReplaceAll As Long = 2
Private Const wdCollapseStart As Long = 1
Private Const wdFormatXMLDocument As Long = 12
Private Const wdDoNotSaveChanges As Long = 0

Public Function RunContractAutomation() As String
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
End Function

Public Function RunContractWorkflow(ByVal wb As Workbook, ByVal templateCode As String, ByVal cfg As Object) As String
    Dim overrideSheet As String
    Dim items As Collection
    Dim keys As Variant
    Dim k As Variant
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

    For Each k In keys
        If cfg.Exists(CStr(k)) Then
            Set templateCfg = cfg(CStr(k))
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
    Next k

    RunContractWorkflow = lastFolder
End Function

Public Function LoadCfgTemplates(ByVal wb As Workbook) As Object
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim cfg As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim code As String
    Dim rowDict As Object

    Set ws = wb.Sheets("CFG_TEMPLATES")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = 1

    For c = 1 To lastCol
        headerMap(Trim$(CellText(ws.Cells(1, c).Value))) = c
    Next c

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = 1

    For r = 2 To lastRow
        code = Trim$(GetCellByHeader(ws, r, headerMap, "TemplateCode"))
        If Len(code) > 0 Then
            Set rowDict = CreateObject("Scripting.Dictionary")
            rowDict.CompareMode = 1
            rowDict("excel_sheet") = GetCellByHeader(ws, r, headerMap, "ExcelSheet")
            rowDict("docx_file") = GetCellByHeader(ws, r, headerMap, "DocxFile")
            rowDict("file_prefix") = GetCellByHeader(ws, r, headerMap, "FilePrefix")
            rowDict("description") = GetCellByHeader(ws, r, headerMap, "Description")
            rowDict("enabled") = ParseEnabled(GetCellByHeader(ws, r, headerMap, "Enabled"))
            cfg(code) = rowDict
        End If
    Next r

    Set LoadCfgTemplates = cfg
End Function

Public Function LoadItems(ByVal wb As Workbook) As Collection
    Dim items As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headers() As String
    Dim r As Long
    Dim c As Long
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
    For c = 1 To lastCol
        headers(c) = NormalizeHeader(CellText(ws.Cells(1, c).Value))
    Next c

    For r = 2 To lastRow
        rowHasData = False
        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = 1

        For c = 1 To lastCol
            item(headers(c)) = CellText(ws.Cells(r, c).Value)
            If Len(Trim$(item(headers(c)))) > 0 Then rowHasData = True
        Next c

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
    Next r

    Set LoadItems = items
End Function

Public Function BuildContext(ByVal wb As Workbook, ByVal sheetName As String) As Object
    Dim ws As Worksheet
    Dim ctx As Object
    Dim lastRow As Long
    Dim r As Long
    Dim k As String
    Dim v As String

    Set ws = wb.Sheets(sheetName)
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.CompareMode = 1

    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, 4).End(xlUp).Row)

    For r = 1 To lastRow
        k = Trim$(CellText(ws.Cells(r, 1).Value))
        If Len(k) > 0 Then
            v = CellText(ws.Cells(r, 4).Value)
            ctx(k) = v
        End If
    Next r

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

Public Function RenderTemplate(ByVal templateCfg As Object, ByVal ctx As Object, ByVal seq As String, ByVal customerName As String, ByVal wb As Workbook) As String
    Dim templatePath As String
    Dim outputRoot As String
    Dim outputFolder As String
    Dim outputPath As String
    Dim wordApp As Object
    Dim doc As Object
    Dim createdWord As Boolean

    templatePath = BuildPath(wb.Path, "Templates", GetDictString(templateCfg, "docx_file"))
    If Dir$(templatePath, vbNormal) = vbNullString Then
        Err.Raise vbObjectError + 513, "RenderTemplate", "Missing template file: " & templatePath
    End If

    outputRoot = BuildPath(wb.Path, "Output")
    EnsureFolderExists outputRoot

    outputFolder = BuildPath(outputRoot, MakeSafeFilename(customerName))
    EnsureFolderExists outputFolder

    outputPath = BuildPath(outputFolder, seq & "_" & GetDictString(templateCfg, "file_prefix") & "_" & MakeSafeFilename(customerName) & ".docx")

    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo 0

    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
        createdWord = True
    End If

    wordApp.Visible = False
    Set doc = wordApp.Documents.Open(templatePath, False, False)

    ApplyContextToDocument doc, ctx
    doc.SaveAs2 outputPath, wdFormatXMLDocument
    doc.Close wdDoNotSaveChanges

    If createdWord Then wordApp.Quit wdDoNotSaveChanges

    Set doc = Nothing
    Set wordApp = Nothing

    RenderTemplate = outputFolder
End Function

Private Sub ApplyContextToDocument(ByVal doc As Object, ByVal ctx As Object)
    Dim items As Collection
    Dim story As Object
    Dim itemValue As Variant

    On Error Resume Next
    Set items = ctx("items")
    On Error GoTo 0

    For Each story In doc.StoryRanges
        ApplyScalarReplacements story, ctx
        If Not items Is Nothing Then ExpandItemsInStory story, items
    Next story

    If Not items Is Nothing Then
        For Each itemValue In items
            ' no-op; forces late binding resolution during compile/import
        Next itemValue
    End If
End Sub

Private Sub ApplyScalarReplacements(ByVal rng As Object, ByVal ctx As Object)
    Dim key As Variant
    Dim valueText As String

    For Each key In ctx.Keys
        If LCase$(CStr(key)) <> "items" Then
            valueText = CellText(ctx(key))
            ReplaceTokenInRange rng, CStr(key), valueText
        End If
    Next key
End Sub

Private Sub ExpandItemsInStory(ByVal rng As Object, ByVal items As Collection)
    Dim tbl As Object
    For Each tbl In rng.Tables
        ExpandItemsInTable tbl, items
    Next tbl
End Sub

Private Sub ExpandItemsInTable(ByVal tbl As Object, ByVal items As Collection)
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long
    Dim rowText As String
    Dim currentRow As Long
    Dim itemIndex As Long

    startRow = 0
    endRow = 0

    For i = 1 To tbl.Rows.Count
        rowText = CleanWordCellText(tbl.Rows(i).Range.Text)
        If InStr(1, rowText, "for item in items", vbTextCompare) > 0 Then startRow = i
        If InStr(1, rowText, "endfor", vbTextCompare) > 0 Then
            endRow = i
            Exit For
        End If
    Next i

    If startRow = 0 Or endRow = 0 Then Exit Sub
    If endRow - startRow <> 2 Then Exit Sub

    currentRow = startRow + 1

    If items.Count = 0 Then
        tbl.Rows(endRow).Delete
        tbl.Rows(currentRow).Delete
        tbl.Rows(startRow).Delete
        Exit Sub
    End If

    For i = 2 To items.Count
        tbl.Rows(currentRow).Select
        tbl.Application.Selection.InsertRowsBelow 1
    Next i

    itemIndex = 1
    Do While itemIndex <= items.Count
        ReplaceItemRow tbl.Rows(currentRow).Range, items(itemIndex), itemIndex
        currentRow = currentRow + 1
        itemIndex = itemIndex + 1
    Loop

    tbl.Rows(currentRow).Delete
    tbl.Rows(startRow).Delete
End Sub

Private Sub ReplaceItemRow(ByVal rng As Object, ByVal item As Object, ByVal itemIndex As Long)
    Dim key As Variant
    ReplaceTokenInRange rng, "loop.index", CStr(itemIndex)

    For Each key In item.Keys
        ReplaceTokenInRange rng, "item." & CStr(key), CellText(item(key))
    Next key
End Sub

Private Sub ReplaceTokenInRange(ByVal rng As Object, ByVal tokenName As String, ByVal replaceText As String)
    ReplaceAllInRange rng, "{{" & tokenName & "}}", replaceText
    ReplaceAllInRange rng, "{{ " & tokenName & " }}", replaceText
    ReplaceAllInRange rng, "{{" & tokenName & " }}", replaceText
    ReplaceAllInRange rng, "{{ " & tokenName & "}}", replaceText

    ReplaceWildcardToken rng, tokenName, replaceText, True, True
    ReplaceWildcardToken rng, tokenName, replaceText, True, False
    ReplaceWildcardToken rng, tokenName, replaceText, False, True
End Sub

Private Sub ReplaceWildcardToken(ByVal rng As Object, ByVal tokenName As String, ByVal replaceText As String, ByVal allowLeftSpaces As Boolean, ByVal allowRightSpaces As Boolean)
    Dim leftPart As String
    Dim rightPart As String
    leftPart = IIf(allowLeftSpaces, "[ ]@", "")
    rightPart = IIf(allowRightSpaces, "[ ]@", "")

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "\{\{" & leftPart & EscapeWordFindText(tokenName) & rightPart & "\}\}"
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub ReplaceAllInRange(ByVal rng As Object, ByVal findText As String, ByVal replaceText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Function EscapeWordFindText(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace$(result, "\\", "\\")
    result = Replace$(result, "{", "\{")
    result = Replace$(result, "}", "\}")
    result = Replace$(result, ".", "\.")
    result = Replace$(result, "(", "\(")
    result = Replace$(result, ")", "\)")
    result = Replace$(result, "[", "\[")
    result = Replace$(result, "]", "\]")
    result = Replace$(result, "-", "\-")
    EscapeWordFindText = result
End Function

Public Function ToNumber(ByVal value As Variant) As Double
    Dim s As String

    If IsError(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function

    If IsNumeric(value) Then
        ToNumber = CDbl(value)
        Exit Function
    End If

    s = Trim$(CStr(value))
    If Len(s) = 0 Then Exit Function

    s = Replace$(s, ChrW$(160), vbNullString)
    s = Replace$(s, " ", vbNullString)

    If InStr(s, ".") > 0 And InStr(s, ",") > 0 Then
        s = Replace$(s, ".", vbNullString)
        s = Replace$(s, ",", ".")
    Else
        If CountOccurrences(s, ".") > 1 Then s = Replace$(s, ".", vbNullString)
        s = Replace$(s, ",", ".")
    End If

    s = KeepNumericChars(s)
    If Len(s) = 0 Or s = "-" Or s = "." Or s = "-." Then Exit Function

    On Error Resume Next
    ToNumber = CDbl(s)
    On Error GoTo 0
End Function

Public Function FormatVN(ByVal n As Variant, Optional ByVal forceDecimals As Boolean = False, Optional ByVal decimals As Variant) As String
    Dim num As Double
    Dim dec As Long
    Dim rounded As Double
    Dim signText As String
    Dim absText As String
    Dim parts() As String
    Dim intPart As String
    Dim fracPart As String
    Dim mask As String

    If IsNull(n) Or IsEmpty(n) Or CellText(n) = vbNullString Then Exit Function

    num = CDbl(n)

    If forceDecimals Then
        If IsMissing(decimals) Or IsEmpty(decimals) Then
            dec = 2
        Else
            dec = CLng(decimals)
        End If
    Else
        If IsMissing(decimals) Or IsEmpty(decimals) Then
            If num = Fix(num) Then
                dec = 0
            Else
                dec = 2
            End If
        Else
            dec = CLng(decimals)
        End If
    End If

    rounded = RoundHalfUpValue(num, dec)
    If rounded < 0 Then signText = "-"

    intPart = Format$(Fix(Abs(rounded)), "#,##0")
    intPart = Replace$(intPart, ",", ".")

    If dec = 0 Then
        FormatVN = signText & intPart
        Exit Function
    End If

    mask = "0." & String$(dec, "0")
    absText = Replace$(Format$(Abs(rounded), mask), ",", ".")
    parts = Split(absText, ".")
    If UBound(parts) >= 1 Then fracPart = parts(1) Else fracPart = String$(dec, "0")

    FormatVN = signText & intPart & "," & fracPart
End Function

Public Function RoundHalfUpValue(ByVal value As Double, Optional ByVal decimals As Long = 0) As Double
    Dim factor As Double
    factor = 10 ^ decimals

    If value >= 0 Then
        RoundHalfUpValue = Int(value * factor + 0.5) / factor
    Else
        RoundHalfUpValue = -Int(Abs(value) * factor + 0.5) / factor
    End If
End Function

Public Function MakeSafeFilename(ByVal text As String) As String
    Dim cleaned As String
    cleaned = RemoveVietnameseDiacritics(Trim$(text))
    cleaned = RegexReplace(cleaned, "[^A-Za-z0-9_\- ]+", "_")
    cleaned = Replace$(cleaned, " ", "_")
    cleaned = Trim$(cleaned)
    Do While InStr(cleaned, "__") > 0
        cleaned = Replace$(cleaned, "__", "_")
    Loop
    If Len(cleaned) = 0 Then cleaned = "contract"
    MakeSafeFilename = cleaned
End Function

Private Function RemoveVietnameseDiacritics(ByVal text As String) As String
    Dim src As Variant
    Dim dst As Variant
    Dim i As Long

    src = Array( _
        "à", "á", "ạ", "ả", "ã", "â", "ầ", "ấ", "ậ", "ẩ", "ẫ", "ă", "ằ", "ắ", "ặ", "ẳ", "ẵ", _
        "è", "é", "ẹ", "ẻ", "ẽ", "ê", "ề", "ế", "ệ", "ể", "ễ", _
        "ì", "í", "ị", "ỉ", "ĩ", _
        "ò", "ó", "ọ", "ỏ", "õ", "ô", "ồ", "ố", "ộ", "ổ", "ỗ", "ơ", "ờ", "ớ", "ợ", "ở", "ỡ", _
        "ù", "ú", "ụ", "ủ", "ũ", "ư", "ừ", "ứ", "ự", "ử", "ữ", _
        "ỳ", "ý", "ỵ", "ỷ", "ỹ", "đ", _
        "À", "Á", "Ạ", "Ả", "Ã", "Â", "Ầ", "Ấ", "Ậ", "Ẩ", "Ẫ", "Ă", "Ằ", "Ắ", "Ặ", "Ẳ", "Ẵ", _
        "È", "É", "Ẹ", "Ẻ", "Ẽ", "Ê", "Ề", "Ế", "Ệ", "Ể", "Ễ", _
        "Ì", "Í", "Ị", "Ỉ", "Ĩ", _
        "Ò", "Ó", "Ọ", "Ỏ", "Õ", "Ô", "Ồ", "Ố", "Ộ", "Ổ", "Ỗ", "Ơ", "Ờ", "Ớ", "Ợ", "Ở", "Ỡ", _
        "Ù", "Ú", "Ụ", "Ủ", "Ũ", "Ư", "Ừ", "Ứ", "Ự", "Ử", "Ữ", _
        "Ỳ", "Ý", "Ỵ", "Ỷ", "Ỹ", "Đ")
    dst = Array( _
        "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", _
        "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", _
        "i", "i", "i", "i", "i", _
        "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", _
        "u", "u", "u", "u", "u", "u", "u", "u", "u", "u", "u", _
        "y", "y", "y", "y", "y", "d", _
        "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", _
        "E", "E", "E", "E", "E", "E", "E", "E", "E", "E", "E", _
        "I", "I", "I", "I", "I", _
        "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", _
        "U", "U", "U", "U", "U", "U", "U", "U", "U", "U", "U", _
        "Y", "Y", "Y", "Y", "Y", "D")

    RemoveVietnameseDiacritics = text
    For i = LBound(src) To UBound(src)
        RemoveVietnameseDiacritics = Replace$(RemoveVietnameseDiacritics, src(i), dst(i))
    Next i
End Function

Private Function NormalizeSequence(ByVal value As String) As String
    Dim num As Double
    If Len(Trim$(value)) = 0 Then
        NormalizeSequence = "00"
        Exit Function
    End If

    num = ToNumber(value)
    NormalizeSequence = Right$("00" & CStr(CLng(num)), 2)
End Function

Private Function NormalizeHeader(ByVal value As String) As String
    Dim s As String
    s = LCase$(Trim$(value))
    s = Replace$(s, " ", "_")
    NormalizeHeader = s
End Function

Private Function ParseEnabled(ByVal value As String) As Boolean
    Select Case UCase$(Trim$(value))
        Case "1", "TRUE", "YES", "Y"
            ParseEnabled = True
    End Select
End Function

Private Function GetCellByHeader(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerMap As Object, ByVal headerName As String) As String
    If headerMap.Exists(headerName) Then
        GetCellByHeader = CellText(ws.Cells(rowNo, CLng(headerMap(headerName))).Value)
    Else
        GetCellByHeader = vbNullString
    End If
End Function

Private Function GetDictString(ByVal dict As Object, ByVal key As String, Optional ByVal defaultValue As String = "") As String
    If dict.Exists(key) Then
        GetDictString = CellText(dict(key))
    Else
        GetDictString = defaultValue
    End If
End Function

Private Function GetDictBoolean(ByVal dict As Object, ByVal key As String) As Boolean
    If dict.Exists(key) Then GetDictBoolean = CBool(dict(key))
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = vbNullString
    Else
        CellText = CStr(value)
    End If
End Function

Private Function BuildPath(ParamArray parts() As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(parts) To UBound(parts)
        If Len(CStr(parts(i))) > 0 Then
            If Len(result) = 0 Then
                result = CStr(parts(i))
            ElseIf Right$(result, 1) = "\" Then
                result = result & CStr(parts(i))
            Else
                result = result & "\" & CStr(parts(i))
            End If
        End If
    Next i
    BuildPath = result
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function KeepNumericChars(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then
            KeepNumericChars = KeepNumericChars & ch
        End If
    Next i
End Function

Private Function CountOccurrences(ByVal text As String, ByVal token As String) As Long
    CountOccurrences = (Len(text) - Len(Replace$(text, token, vbNullString))) / Len(token)
End Function

Private Function RegexReplace(ByVal text As String, ByVal pattern As String, ByVal replaceWith As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = pattern
    RegexReplace = re.Replace(text, replaceWith)
End Function

Private Function CleanWordCellText(ByVal text As String) As String
    CleanWordCellText = Replace$(Replace$(text, Chr$(13), vbNullString), Chr$(7), vbNullString)
End Function


