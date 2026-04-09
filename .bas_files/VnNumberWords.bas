Attribute VB_Name = "VnNumberWords"
Option Explicit

Public Function number_to_words(ByVal value As Variant, Optional ByVal decimal_places As Long = 0, Optional ByVal use_commas As Boolean = False) As String
    number_to_words = NumberToWords(value, decimal_places, use_commas)
End Function

Public Function vnd_to_words(ByVal amount As Variant, Optional ByVal use_commas As Boolean = False, Optional ByVal append_chan As Boolean = True) As String
    vnd_to_words = VndToWords(amount, use_commas, append_chan)
End Function

Public Function NumberToWords(ByVal value As Variant, Optional ByVal decimalPlaces As Long = 0, Optional ByVal useCommas As Boolean = False) As String
    Dim numberValue As Double
    Dim normalizedText As String
    Dim decimalText As String
    Dim result As String

    If decimalPlaces < 0 Then
        NumberToWords = vbNullString
        Exit Function
    End If

    If Not TryNormalizeNumber(value, numberValue, normalizedText) Then
        NumberToWords = vbNullString
        Exit Function
    End If

    result = ReadIntegerPart(Fix(Abs(numberValue)), useCommas)

    If decimalPlaces > 0 Then
        decimalText = TruncateDecimalText(normalizedText, decimalPlaces)
        result = result & " " & WordPhay() & " " & ReadDecimalPart(decimalText)
    End If

    If numberValue < 0 Then
        result = WordAm() & " " & result
    End If

    NumberToWords = UpperFirst(result)
End Function

Public Function VndToWords(ByVal amount As Variant, Optional ByVal useCommas As Boolean = False, Optional ByVal appendChan As Boolean = True) As String
    Dim numberValue As Double
    Dim normalizedText As String
    Dim result As String
    Dim hasDecimal As Boolean

    If Not TryNormalizeNumber(amount, numberValue, normalizedText) Then
        VndToWords = vbNullString
        Exit Function
    End If

    hasDecimal = HasNonZeroDecimal(normalizedText)
    result = ReadIntegerPart(Fix(Abs(numberValue)), useCommas)

    If appendChan And Not hasDecimal Then
        result = result & " " & WordDong() & " " & WordChan()
    Else
        result = result & " " & WordDong()
    End If

    If numberValue < 0 Then
        result = WordAm() & " " & result
    End If

    VndToWords = UpperFirst(result)
End Function

Private Function TryNormalizeNumber(ByVal value As Variant, ByRef numberValue As Double, ByRef normalizedText As String) As Boolean
    Dim textValue As String
    Dim decimalSep As String
    Dim thousandSep As String
    Dim lastComma As Long
    Dim lastDot As Long
    Dim commaCount As Long
    Dim dotCount As Long
    Dim currentDecimalSep As String
    Dim currentThousandSep As String

    On Error GoTo Fail

    textValue = Trim$(CStr(value))
    textValue = Replace(textValue, ChrW$(160), "")
    textValue = Replace(textValue, " ", "")

    If Len(textValue) = 0 Then GoTo Fail

    currentDecimalSep = Application.International(xlDecimalSeparator)
    currentThousandSep = Application.International(xlThousandsSeparator)

    lastComma = InStrRev(textValue, ",")
    lastDot = InStrRev(textValue, ".")
    commaCount = Len(textValue) - Len(Replace(textValue, ",", ""))
    dotCount = Len(textValue) - Len(Replace(textValue, ".", ""))

    decimalSep = vbNullString
    thousandSep = vbNullString

    If lastComma > 0 And lastDot > 0 Then
        If lastComma > lastDot Then
            decimalSep = ","
            thousandSep = "."
        Else
            decimalSep = "."
            thousandSep = ","
        End If
    ElseIf lastComma > 0 Then
        ResolveSingleSeparator textValue, ",", commaCount, decimalSep, thousandSep
    ElseIf lastDot > 0 Then
        ResolveSingleSeparator textValue, ".", dotCount, decimalSep, thousandSep
    End If

    If Len(thousandSep) > 0 Then textValue = Replace(textValue, thousandSep, "")
    If Len(currentThousandSep) > 0 And currentThousandSep <> decimalSep Then
        textValue = Replace(textValue, currentThousandSep, "")
    End If

    If Len(decimalSep) > 0 And decimalSep <> currentDecimalSep Then
        textValue = Replace(textValue, decimalSep, currentDecimalSep)
    End If

    If Not IsNumeric(textValue) Then GoTo Fail

    numberValue = CDbl(textValue)
    normalizedText = Replace(textValue, currentDecimalSep, ".")
    TryNormalizeNumber = True
    Exit Function

Fail:
    numberValue = 0#
    normalizedText = vbNullString
    TryNormalizeNumber = False
End Function

Private Sub ResolveSingleSeparator(ByVal textValue As String, ByVal sep As String, ByVal sepCount As Long, ByRef decimalSep As String, ByRef thousandSep As String)
    Dim lastPos As Long
    Dim digitsRight As Long

    lastPos = InStrRev(textValue, sep)
    digitsRight = Len(textValue) - lastPos

    If sepCount > 1 Then
        thousandSep = sep
        Exit Sub
    End If

    If digitsRight = 3 Then
        thousandSep = sep
    Else
        decimalSep = sep
    End If
End Sub

Private Function ReadIntegerPart(ByVal numberValue As Double, Optional ByVal useCommas As Boolean = False) As String
    Dim groups() As String
    Dim groupIndex As Long
    Dim groupValue As Long
    Dim highestGroup As Long
    Dim text As String
    Dim result As String
    Dim separator As String

    If numberValue = 0 Then
        ReadIntegerPart = DigitWord(0)
        Exit Function
    End If

    groups = SplitThousands(CStr(Fix(numberValue)))
    highestGroup = HighestNonZeroGroup(groups)
    separator = " "
    If useCommas Then separator = ", "

    For groupIndex = UBound(groups) To LBound(groups) Step -1
        groupValue = CLng(groups(groupIndex))
        If groupValue <> 0 Then
            text = ReadThreeDigits(groupValue, groupIndex < highestGroup)
            If Len(GroupUnit(groupIndex)) > 0 Then
                text = text & " " & GroupUnit(groupIndex)
            End If

            If Len(result) = 0 Then
                result = text
            Else
                result = result & separator & text
            End If
        End If
    Next groupIndex

    ReadIntegerPart = result
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

Private Function ReadDecimalPart(ByVal decimalText As String) As String
    Dim i As Long
    Dim text As String

    For i = 1 To Len(decimalText)
        AddWord text, DigitWord(CLng(Mid$(decimalText, i, 1)))
    Next i

    ReadDecimalPart = text
End Function

Private Function TruncateDecimalText(ByVal normalizedText As String, ByVal decimalPlaces As Long) As String
    Dim p As Long
    Dim decimalText As String

    p = InStr(1, normalizedText, ".", vbBinaryCompare)
    If p = 0 Then
        TruncateDecimalText = String$(decimalPlaces, "0")
        Exit Function
    End If

    decimalText = Mid$(normalizedText, p + 1)
    decimalText = Left$(decimalText, decimalPlaces)

    If Len(decimalText) < decimalPlaces Then
        decimalText = decimalText & String$(decimalPlaces - Len(decimalText), "0")
    End If

    TruncateDecimalText = decimalText
End Function

Private Function HasNonZeroDecimal(ByVal normalizedText As String) As Boolean
    Dim p As Long
    Dim decimalText As String

    p = InStr(1, normalizedText, ".", vbBinaryCompare)
    If p = 0 Then Exit Function

    decimalText = Mid$(normalizedText, p + 1)
    decimalText = Replace(decimalText, "0", "")
    HasNonZeroDecimal = (Len(decimalText) > 0)
End Function

Private Function SplitThousands(ByVal textValue As String) As String()
    Dim groups() As String
    Dim partText As String
    Dim i As Long

    i = -1
    Do While Len(textValue) > 0
        i = i + 1
        ReDim Preserve groups(0 To i)
        If Len(textValue) > 3 Then
            partText = Right$(textValue, 3)
            textValue = Left$(textValue, Len(textValue) - 3)
        Else
            partText = textValue
            textValue = vbNullString
        End If
        groups(i) = partText
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
    If unitValue = 1 And tenValue > 1 Then
        UnitWord = VnText("6D;1ED1;74")
    ElseIf unitValue = 5 And tenValue >= 1 Then
        UnitWord = VnText("6C;103;6D")
    Else
        UnitWord = DigitWord(unitValue)
    End If
End Function

Private Function GroupUnit(ByVal groupIndex As Long) As String
    Select Case groupIndex
        Case 0: GroupUnit = vbNullString
        Case 1: GroupUnit = VnText("6E;67;68;EC;6E")
        Case 2: GroupUnit = VnText("74;72;69;1EC7;75")
        Case 3: GroupUnit = VnText("74;1EF7")
        Case 4: GroupUnit = VnText("6E;67;68;EC;6E;20;74;1EF7")
        Case Else: GroupUnit = vbNullString
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

Private Function WordPhay() As String
    WordPhay = VnText("70;68;1EA9;79")
End Function

Private Function WordAm() As String
    WordAm = VnText("E2;6D")
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
