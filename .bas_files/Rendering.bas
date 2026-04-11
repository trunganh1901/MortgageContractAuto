Attribute VB_Name = "Rendering"
Option Explicit

Public Function RenderTemplate(ByVal templateCfg As Object, ByVal ctx As Object, ByVal wb As Workbook) As String
    Dim templatePath As String
    Dim outputRoot As String
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
    outputPath = BuildAvailableOutputPath(outputRoot, GetDictString(templateCfg, "file_prefix", "document"))

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

    RenderTemplate = outputPath
End Function

Private Sub ApplyContextToDocument(ByVal doc As Object, ByVal ctx As Object)
    Dim story As Object

    For Each story In doc.StoryRanges
        ApplyScalarReplacements story, ctx
    Next story
End Sub

Private Sub ApplyScalarReplacements(ByVal rng As Object, ByVal ctx As Object)
    Dim key As Variant
    Dim valueText As String

    For Each key In ctx.Keys
        valueText = CellText(ctx(key))
        ReplaceTokenInRange rng, CStr(key), valueText
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

Private Function BuildAvailableOutputPath(ByVal outputRoot As String, ByVal filePrefix As String) As String
    Dim baseName As String
    Dim versionNo As Long
    Dim candidatePath As String

    baseName = MakeSafeFilename(filePrefix)
    If Len(baseName) = 0 Then baseName = "document"

    versionNo = 1
    Do
        candidatePath = BuildPath(outputRoot, baseName & "_v" & CStr(versionNo) & ".docx")
        If Dir$(candidatePath, vbNormal) = vbNullString Then
            BuildAvailableOutputPath = candidatePath
            Exit Function
        End If
        versionNo = versionNo + 1
    Loop
End Function
