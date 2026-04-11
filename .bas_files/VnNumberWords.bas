Attribute VB_Name = "VnNumberWords"
Option Explicit

Private Const AUTO_DECIMAL_PLACES As Long = -1

Public Function number_to_words(ByVal value As Variant, Optional ByVal decimal_places As Long = AUTO_DECIMAL_PLACES, Optional ByVal use_commas As Boolean = False) As String
    number_to_words = NumberToWords(value, decimal_places, use_commas)
End Function

Public Function vnd_to_words(ByVal amount As Variant, Optional ByVal use_commas As Boolean = False, Optional ByVal append_chan As Boolean = True) As String
    vnd_to_words = VndToWords(amount, use_commas, append_chan)
End Function

Public Function NumberToWords(ByVal value As Variant, Optional ByVal decimalPlaces As Long = AUTO_DECIMAL_PLACES, Optional ByVal useCommas As Boolean = False) As String
    Dim negativeValue As Boolean
    Dim integerDigits As String
    Dim decimalDigits As String
    Dim result As String

    If decimalPlaces < AUTO_DECIMAL_PLACES Then Exit Function
    If Not ParseInputNumber(value, negativeValue, integerDigits, decimalDigits) Then Exit Function

    result = ReadIntegerDigits(integerDigits, useCommas)
    decimalDigits = NormalizeDecimalDigits(decimalDigits, decimalPlaces)

    If Len(decimalDigits) > 0 Then
        result = result & " " & WordPhay() & " " & ReadDecimalDigits(decimalDigits)
    End If

    If negativeValue And result <> DigitWord(0) Then
        result = WordAm() & " " & result
    End If

    NumberToWords = UpperFirst(result)
End Function

Public Function VndToWords(ByVal amount As Variant, Optional ByVal useCommas As Boolean = False, Optional ByVal appendChan As Boolean = True) As String
    Dim negativeValue As Boolean
    Dim integerDigits As String
    Dim decimalDigits As String
    Dim result As String

    If Not ParseInputNumber(amount, negativeValue, integerDigits, decimalDigits) Then Exit Function

    result = ReadIntegerDigits(integerDigits, useCommas)

    If Len(decimalDigits) > 0 Then
        result = result & " " & WordPhay() & " " & ReadDecimalDigits(decimalDigits)
        result = result & " " & WordDong()
    ElseIf appendChan Then
        result = result & " " & WordDong() & " " & WordChan()
    Else
        result = result & " " & WordDong()
    End If

    If negativeValue And result <> DigitWord(0) & " " & WordDong() Then
        result = WordAm() & " " & result
    End If

    VndToWords = UpperFirst(result)
End Function

Private Function ParseInputNumber(ByVal value As Variant, ByRef negativeValue As Boolean, ByRef integerDigits As String, ByRef decimalDigits As String) As Boolean
    Dim rawText As String

    If TryParseNumericInput(value, negativeValue, integerDigits, decimalDigits) Then
        ParseInputNumber = True
        Exit Function
    End If

    rawText = InputText(value)
    If Len(rawText) = 0 Then Exit Function

    ParseInputNumber = ParseNumberText(rawText, negativeValue, integerDigits, decimalDigits)
End Function

Private Function TryParseNumericInput(ByVal value As Variant, ByRef negativeValue As Boolean, ByRef integerDigits As String, ByRef decimalDigits As String) As Boolean
    If IsObject(value) Then
        If TypeName(value) <> "Range" Then Exit Function
        TryParseNumericInput = ParseNumericVariant(value.Value2, negativeValue, integerDigits, decimalDigits)
        Exit Function
    End If

    TryParseNumericInput = ParseNumericVariant(value, negativeValue, integerDigits, decimalDigits)
End Function

Private Function InputText(ByVal value As Variant) As String
    On Error Resume Next

    If IsObject(value) Then
        If TypeName(value) = "Range" Then
            InputText = CStr(value.Value2)
        End If
    Else
        InputText = CStr(value)
    End If
End Function

Private Function ParseNumericVariant(ByVal value As Variant, ByRef negativeValue As Boolean, ByRef integerDigits As String, ByRef decimalDigits As String) As Boolean
    Dim rawText As String

    If VarType(value) = vbString Then Exit Function
    If Not IsNumeric(value) Then Exit Function

    rawText = NumericValueToText(CDbl(value))
    If Len(rawText) = 0 Then Exit Function

    ParseNumericVariant = ParseNumberText(rawText, negativeValue, integerDigits, decimalDigits)
End Function

Private Function NumericValueToText(ByVal numberValue As Double) As String
    Dim currentDecimalSep As String
    Dim formattedText As String

    currentDecimalSep = Application.International(xlDecimalSeparator)
    formattedText = Format$(Round(numberValue, 12), "0.############")

    If Len(formattedText) = 0 Then Exit Function

    If currentDecimalSep <> "." Then
        formattedText = Replace(formattedText, ".", currentDecimalSep)
    End If

    NumericValueToText = formattedText
End Function

Private Function ParseNumberText(ByVal rawText As String, ByRef negativeValue As Boolean, ByRef integerDigits As String, ByRef decimalDigits As String) As Boolean
    Dim textValue As String
    Dim decimalSep As String
    Dim thousandSep As String
    Dim currentDecimalSep As String
    Dim currentThousandSep As String
    Dim lastComma As Long
    Dim lastDot As Long
    Dim commaCount As Long
    Dim dotCount As Long
    Dim splitPos As Long

    textValue = SanitizeNumberText(rawText)
    If Len(textValue) = 0 Then Exit Function

    If Not StripSign(textValue, negativeValue) Then Exit Function

    currentDecimalSep = Application.International(xlDecimalSeparator)
    currentThousandSep = Application.International(xlThousandsSeparator)

    lastComma = InStrRev(textValue, ",")
    lastDot = InStrRev(textValue, ".")
    commaCount = Len(textValue) - Len(Replace(textValue, ",", ""))
    dotCount = Len(textValue) - Len(Replace(textValue, ".", ""))

    ResolveSeparators textValue, lastComma, lastDot, commaCount, dotCount, currentDecimalSep, currentThousandSep, decimalSep, thousandSep

    If Len(thousandSep) > 0 Then
        textValue = Replace(textValue, thousandSep, "")
    End If

    If Len(decimalSep) > 0 Then
        splitPos = InStrRev(textValue, decimalSep)
        integerDigits = Left$(textValue, splitPos - 1)
        decimalDigits = Mid$(textValue, splitPos + 1)
    Else
        integerDigits = textValue
        decimalDigits = vbNullString
    End If

    integerDigits = KeepDigitsOnly(integerDigits)
    decimalDigits = KeepDigitsOnly(decimalDigits)

    If Len(integerDigits) = 0 And Len(decimalDigits) = 0 Then Exit Function
    If Len(integerDigits) = 0 Then integerDigits = "0"

    integerDigits = TrimLeadingZeros(integerDigits)
    ParseNumberText = True
End Function

Private Function SanitizeNumberText(ByVal rawText As String) As String
    Dim textValue As String

    textValue = Trim$(rawText)
    textValue = Replace(textValue, ChrW$(160), "")
    textValue = Replace(textValue, " ", "")
    textValue = Replace(textValue, vbTab, "")

    SanitizeNumberText = textValue
End Function

Private Function StripSign(ByRef textValue As String, ByRef negativeValue As Boolean) As Boolean
    If Len(textValue) = 0 Then Exit Function

    Select Case Left$(textValue, 1)
        Case "+"
            textValue = Mid$(textValue, 2)
        Case "-"
            negativeValue = True
            textValue = Mid$(textValue, 2)
    End Select

    StripSign = (Len(textValue) > 0)
End Function

Private Sub ResolveSeparators(ByVal textValue As String, ByVal lastComma As Long, ByVal lastDot As Long, ByVal commaCount As Long, ByVal dotCount As Long, ByVal currentDecimalSep As String, ByVal currentThousandSep As String, ByRef decimalSep As String, ByRef thousandSep As String)
    If lastComma > 0 And lastDot > 0 Then
        If lastComma > lastDot Then
            decimalSep = ","
            thousandSep = "."
        Else
            decimalSep = "."
            thousandSep = ","
        End If
        Exit Sub
    End If

    If lastComma > 0 Then
        ResolveSingleSeparator textValue, ",", commaCount, currentDecimalSep, currentThousandSep, decimalSep, thousandSep
    ElseIf lastDot > 0 Then
        ResolveSingleSeparator textValue, ".", dotCount, currentDecimalSep, currentThousandSep, decimalSep, thousandSep
    End If
End Sub

Private Sub ResolveSingleSeparator(ByVal textValue As String, ByVal sep As String, ByVal sepCount As Long, ByVal currentDecimalSep As String, ByVal currentThousandSep As String, ByRef decimalSep As String, ByRef thousandSep As String)
    Dim lastPos As Long
    Dim digitsRight As Long

    If sep = currentDecimalSep And sep <> currentThousandSep Then
        decimalSep = sep
        Exit Sub
    End If

    If sep = currentThousandSep And sep <> currentDecimalSep Then
        thousandSep = sep
        Exit Sub
    End If

    lastPos = InStrRev(textValue, sep)
    digitsRight = Len(textValue) - lastPos

    If sepCount > 1 Then
        If digitsRight = 3 Then
            thousandSep = sep
        Else
            decimalSep = sep
        End If
    ElseIf digitsRight = 3 Then
        thousandSep = sep
    Else
        decimalSep = sep
    End If
End Sub

Private Function NormalizeDecimalDigits(ByVal decimalDigits As String, ByVal decimalPlaces As Long) As String
    If decimalPlaces = 0 Then Exit Function

    If decimalPlaces > 0 Then
        decimalDigits = Left$(decimalDigits & String$(decimalPlaces, "0"), decimalPlaces)
    End If

    NormalizeDecimalDigits = decimalDigits
End Function

Private Function ReadIntegerDigits(ByVal integerDigits As String, Optional ByVal useCommas As Boolean = False) As String
    Dim groups() As String
    Dim highestGroup As Long
    Dim groupIndex As Long
    Dim groupText As String
    Dim result As String
    Dim separator As String

    integerDigits = TrimLeadingZeros(KeepDigitsOnly(integerDigits))
    If Len(integerDigits) = 0 Then integerDigits = "0"

    If integerDigits = "0" Then
        ReadIntegerDigits = DigitWord(0)
        Exit Function
    End If

    groups = SplitThousands(integerDigits)
    highestGroup = HighestNonZeroGroup(groups)
    separator = IIf(useCommas, ", ", " ")

    For groupIndex = UBound(groups) To LBound(groups) Step -1
        If CLng(groups(groupIndex)) <> 0 Then
            groupText = ReadThreeDigits(CLng(groups(groupIndex)), groupIndex < highestGroup)

            If Len(GroupUnit(groupIndex)) > 0 Then
                groupText = groupText & " " & GroupUnit(groupIndex)
            End If

            If Len(result) = 0 Then
                result = groupText
            Else
                result = result & separator & groupText
            End If
        End If
    Next groupIndex

    ReadIntegerDigits = result
End Function

Private Function ReadThreeDigits(ByVal groupValue As Long, ByVal fullRead As Boolean) As String
    Dim hundreds As Long
    Dim tens As Long
    Dim units As Long
    Dim text As String

    hundreds = groupValue \ 100
    tens = (groupValue Mod 100) \ 10
    units = groupValue Mod 10

    If hundreds > 0 Then
        AddWord text, DigitWord(hundreds)
        AddWord text, WordTram()
    ElseIf fullRead And groupValue > 0 Then
        AddWord text, DigitWord(0)
        AddWord text, WordTram()
    End If

    Select Case tens
        Case 0
            If units > 0 Then
                If hundreds > 0 Or fullRead Then AddWord text, WordLinh()
                AddWord text, DigitWord(units)
            End If
        Case 1
            AddWord text, WordMuoi10()
            If units > 0 Then AddWord text, UnitWord(units, tens)
        Case Else
            AddWord text, DigitWord(tens)
            AddWord text, WordMuoi()
            If units > 0 Then AddWord text, UnitWord(units, tens)
    End Select

    ReadThreeDigits = text
End Function

Private Function ReadDecimalDigits(ByVal decimalDigits As String) As String
    If Len(decimalDigits) = 0 Then Exit Function

    If Left$(decimalDigits, 1) = "0" And Len(decimalDigits) > 1 Then
        ReadDecimalDigits = ReadDigitByDigit(decimalDigits)
    Else
        ReadDecimalDigits = ReadIntegerDigits(decimalDigits, False)
    End If
End Function

Private Function ReadDigitByDigit(ByVal digitText As String) As String
    Dim i As Long
    Dim text As String

    For i = 1 To Len(digitText)
        AddWord text, DigitWord(CLng(Mid$(digitText, i, 1)))
    Next i

    ReadDigitByDigit = text
End Function

Private Function SplitThousands(ByVal textValue As String) As String()
    Dim groups() As String
    Dim index As Long
    Dim partText As String

    index = -1
    Do While Len(textValue) > 0
        index = index + 1
        ReDim Preserve groups(0 To index)

        If Len(textValue) > 3 Then
            partText = Right$(textValue, 3)
            textValue = Left$(textValue, Len(textValue) - 3)
        Else
            partText = textValue
            textValue = vbNullString
        End If

        groups(index) = partText
    Loop

    SplitThousands = groups
End Function

Private Function HighestNonZeroGroup(ByRef groups() As String) As Long
    Dim i As Long

    For i = UBound(groups) To LBound(groups) Step -1
        If CLng(groups(i)) <> 0 Then
            HighestNonZeroGroup = i
            Exit Function
        End If
    Next i
End Function

Private Function KeepDigitsOnly(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch >= "0" And ch <= "9" Then
            KeepDigitsOnly = KeepDigitsOnly & ch
        End If
    Next i
End Function

Private Function TrimLeadingZeros(ByVal textValue As String) As String
    Do While Len(textValue) > 1 And Left$(textValue, 1) = "0"
        textValue = Mid$(textValue, 2)
    Loop

    TrimLeadingZeros = textValue
End Function

Private Function GroupUnit(ByVal groupIndex As Long) As String
    Select Case groupIndex
        Case 0: GroupUnit = vbNullString
        Case 1: GroupUnit = VnText("6E;67;68;EC;6E")
        Case 2: GroupUnit = VnText("74;72;69;1EC7;75")
        Case 3: GroupUnit = VnText("74;1EF7")
        Case 4: GroupUnit = VnText("6E;67;68;EC;6E;20;74;1EF7")
        Case 5: GroupUnit = VnText("74;72;69;1EC7;75;20;74;1EF7")
        Case 6: GroupUnit = VnText("74;1EF7;20;74;1EF7")
        Case Else: GroupUnit = vbNullString
    End Select
End Function

Private Function DigitWord(ByVal digitValue As Long) As String
    Select Case digitValue
        Case 0: DigitWord = VnText("6B;68;F4;6E;67")
        Case 1: DigitWord = VnText("6D;1ED9;74")
        Case 2: DigitWord = "hai"
        Case 3: DigitWord = "ba"
        Case 4: DigitWord = VnText("62;1ED1;6E")
        Case 5: DigitWord = VnText("6E;103;6D")
        Case 6: DigitWord = VnText("73;E1;75")
        Case 7: DigitWord = VnText("62;1EA3;79")
        Case 8: DigitWord = VnText("74;E1;6D")
        Case 9: DigitWord = VnText("63;68;ED;6E")
    End Select
End Function

Private Function UnitWord(ByVal unitValue As Long, ByVal tenValue As Long) As String
    Select Case unitValue
        Case 1
            If tenValue > 1 Then
                UnitWord = VnText("6D;1ED1;74")
            Else
                UnitWord = VnText("6D;1ED9;74")
            End If
        Case 4
            UnitWord = VnText("62;1ED1;6E")
        Case 5
            If tenValue >= 1 Then
                UnitWord = VnText("6C;103;6D")
            Else
                UnitWord = VnText("6E;103;6D")
            End If
        Case Else
            UnitWord = DigitWord(unitValue)
    End Select
End Function

Private Sub AddWord(ByRef text As String, ByVal wordText As String)
    If Len(wordText) = 0 Then Exit Sub

    If Len(text) = 0 Then
        text = wordText
    Else
        text = text & " " & wordText
    End If
End Sub

Private Function UpperFirst(ByVal textValue As String) As String
    If Len(textValue) = 0 Then Exit Function
    UpperFirst = UCase$(Left$(textValue, 1)) & Mid$(textValue, 2)
End Function

Private Function WordAm() As String
    WordAm = VnText("E2;6D")
End Function

Private Function WordPhay() As String
    WordPhay = VnText("70;68;1EA9;79")
End Function

Private Function WordDong() As String
    WordDong = VnText("111;1ED3;6E;67")
End Function

Private Function WordChan() As String
    WordChan = VnText("63;68;1EB5;6E")
End Function

Private Function WordTram() As String
    WordTram = VnText("74;72;103;6D")
End Function

Private Function WordLinh() As String
    WordLinh = "linh"
End Function

Private Function WordMuoi10() As String
    WordMuoi10 = VnText("6D;1B0;1EDD;69")
End Function

Private Function WordMuoi() As String
    WordMuoi = VnText("6D;1B0;1A1;69")
End Function

Private Function VnText(ByVal hexCodes As String) As String
    Dim parts() As String
    Dim i As Long

    parts = Split(hexCodes, ";")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            VnText = VnText & ChrW$(CLng("&H" & parts(i)))
        End If
    Next i
End Function
