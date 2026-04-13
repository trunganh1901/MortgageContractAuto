Attribute VB_Name = "modShared"
Option Explicit

Public Const DEFAULT_VAT_RATE As Double = 0.08
Public Const wdFindContinue As Long = 1
Public Const wdReplaceAll As Long = 2
Public Const wdFormatXMLDocument As Long = 12
Public Const wdDoNotSaveChanges As Long = 0
Public Const wdCollapseEnd As Long = 0

Public Function ToNumber(ByVal value As Variant) As Double
    Dim textValue As String

    If IsError(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function

    If IsNumeric(value) Then
        ToNumber = CDbl(value)
        Exit Function
    End If

    textValue = Trim$(CStr(value))
    If Len(textValue) = 0 Then Exit Function

    textValue = Replace$(textValue, ChrW$(160), vbNullString)
    textValue = Replace$(textValue, " ", vbNullString)

    If InStr(textValue, ".") > 0 And InStr(textValue, ",") > 0 Then
        textValue = Replace$(textValue, ".", vbNullString)
        textValue = Replace$(textValue, ",", ".")
    Else
        If CountOccurrences(textValue, ".") > 1 Then textValue = Replace$(textValue, ".", vbNullString)
        textValue = Replace$(textValue, ",", ".")
    End If

    textValue = KeepNumericChars(textValue)
    If Len(textValue) = 0 Or textValue = "-" Or textValue = "." Or textValue = "-." Then Exit Function

    On Error Resume Next
    ToNumber = CDbl(textValue)
    On Error GoTo 0
End Function

Public Function FormatVN(ByVal n As Variant, Optional ByVal forceDecimals As Boolean = False, Optional ByVal decimals As Variant) As String
    Dim num As Double
    Dim decimalCount As Long
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
            decimalCount = 2
        Else
            decimalCount = CLng(decimals)
        End If
    Else
        If IsMissing(decimals) Or IsEmpty(decimals) Then
            If num = Fix(num) Then
                decimalCount = 0
            Else
                decimalCount = 2
            End If
        Else
            decimalCount = CLng(decimals)
        End If
    End If

    rounded = RoundHalfUpValue(num, decimalCount)
    If rounded < 0 Then signText = "-"

    intPart = Format$(Fix(Abs(rounded)), "#,##0")
    intPart = Replace$(intPart, ",", ".")

    If decimalCount = 0 Then
        FormatVN = signText & intPart
        Exit Function
    End If

    mask = "0." & String$(decimalCount, "0")
    absText = Replace$(Format$(Abs(rounded), mask), ",", ".")
    parts = Split(absText, ".")
    If UBound(parts) >= 1 Then fracPart = parts(1) Else fracPart = String$(decimalCount, "0")

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

Public Function NormalizeSequence(ByVal value As String) As String
    Dim num As Double

    If Len(Trim$(value)) = 0 Then
        NormalizeSequence = "00"
        Exit Function
    End If

    num = ToNumber(value)
    NormalizeSequence = Right$("00" & CStr(CLng(num)), 2)
End Function

Public Function GetDictString(ByVal dict As Object, ByVal key As String, Optional ByVal defaultValue As String = "") As String
    If dict.Exists(key) Then
        GetDictString = CellText(dict(key))
    Else
        GetDictString = defaultValue
    End If
End Function

Public Function GetDictBoolean(ByVal dict As Object, ByVal key As String) As Boolean
    If dict.Exists(key) Then GetDictBoolean = CBool(dict(key))
End Function

Public Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = vbNullString
    Else
        CellText = CStr(value)
    End If
End Function

Public Function NormalizeLineBreaks(ByVal text As String, Optional ByVal lineBreak As String = vbLf) As String
    Dim normalized As String

    normalized = text
    normalized = Replace$(normalized, vbCrLf, vbLf)
    normalized = Replace$(normalized, vbCr, vbLf)

    If lineBreak <> vbLf Then
        normalized = Replace$(normalized, vbLf, lineBreak)
    End If

    NormalizeLineBreaks = normalized
End Function

Public Function ExcelCellText(ByVal value As Variant) As String
    ExcelCellText = NormalizeLineBreaks(CellText(value), vbLf)
End Function

Public Function WordReplaceText(ByVal value As Variant) As String
    WordReplaceText = NormalizeLineBreaks(CellText(value), vbCr)
End Function

Public Function BuildPath(ParamArray parts() As Variant) As String
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

Public Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Public Function CleanWordCellText(ByVal text As String) As String
    CleanWordCellText = Replace$(Replace$(text, Chr$(13), vbNullString), Chr$(7), vbNullString)
End Function

Public Function EscapeWordFindText(ByVal text As String) As String
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

Private Function RemoveVietnameseDiacritics(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long
    Dim mappedChar As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        codePoint = AscW(ch)
        If codePoint < 0 Then codePoint = codePoint + 65536

        mappedChar = VietnameseBaseLetter(codePoint)
        If Len(mappedChar) = 0 Then mappedChar = ch

        RemoveVietnameseDiacritics = RemoveVietnameseDiacritics & mappedChar
    Next i
End Function

Private Function VietnameseBaseLetter(ByVal codePoint As Long) As String
    Select Case codePoint
        Case &HE0, &HE1, &H1EA1, &H1EA3, &HE3, &HE2, &H1EA7, &H1EA5, &H1EAD, &H1EA9, &H1EAB, &H103, &H1EB1, &H1EAF, &H1EB7, &H1EB3, &H1EB5
            VietnameseBaseLetter = "a"
        Case &HE8, &HE9, &H1EB9, &H1EBB, &H1EBD, &HEA, &H1EC1, &H1EBF, &H1EC7, &H1EC3, &H1EC5
            VietnameseBaseLetter = "e"
        Case &HEC, &HED, &H1ECB, &H1EC9, &H129
            VietnameseBaseLetter = "i"
        Case &HF2, &HF3, &H1ECD, &H1ECF, &HF5, &HF4, &H1ED3, &H1ED1, &H1ED9, &H1ED5, &H1ED7, &H1A1, &H1EDD, &H1EDB, &H1EE3, &H1EDF, &H1EE1
            VietnameseBaseLetter = "o"
        Case &HF9, &HFA, &H1EE5, &H1EE7, &H169, &H1B0, &H1EEB, &H1EE9, &H1EF1, &H1EED, &H1EEF
            VietnameseBaseLetter = "u"
        Case &H1EF3, &HFD, &H1EF5, &H1EF7, &H1EF9
            VietnameseBaseLetter = "y"
        Case &H111
            VietnameseBaseLetter = "d"
        Case &HC0, &HC1, &H1EA0, &H1EA2, &HC3, &HC2, &H1EA6, &H1EA4, &H1EAC, &H1EA8, &H1EAA, &H102, &H1EB0, &H1EAE, &H1EB6, &H1EB2, &H1EB4
            VietnameseBaseLetter = "A"
        Case &HC8, &HC9, &H1EB8, &H1EBA, &H1EBC, &HCA, &H1EC0, &H1EBE, &H1EC6, &H1EC2, &H1EC4
            VietnameseBaseLetter = "E"
        Case &HCC, &HCD, &H1ECA, &H1EC8, &H128
            VietnameseBaseLetter = "I"
        Case &HD2, &HD3, &H1ECC, &H1ECE, &HD5, &HD4, &H1ED2, &H1ED0, &H1ED8, &H1ED4, &H1ED6, &H1A0, &H1EDC, &H1EDA, &H1EE2, &H1EDE, &H1EE0
            VietnameseBaseLetter = "O"
        Case &HD9, &HDA, &H1EE4, &H1EE6, &H168, &H1AF, &H1EEA, &H1EE8, &H1EF0, &H1EEC, &H1EEE
            VietnameseBaseLetter = "U"
        Case &H1EF2, &HDD, &H1EF4, &H1EF6, &H1EF8
            VietnameseBaseLetter = "Y"
        Case &H110
            VietnameseBaseLetter = "D"
    End Select
End Function

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
