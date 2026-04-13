Attribute VB_Name = "Rendering"
Option Explicit

Public Function RenderTemplate(ByVal templateCfg As Object, ByVal ctx As Object, ByVal wb As Workbook) As String
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

    Dim outputBase As String
    Dim filePrefix As String

    outputRoot = BuildPath(wb.Path, "Output")
    outputFolder = BuildStructuredFolder(outputRoot, ctx, GetDictString(templateCfg, "template_code", "document"))
    EnsureFolderTreeExists outputFolder
    outputBase = GetDictString(ctx, "CIF", "") & "_" & GetDictString(ctx, "NAME", "")
    filePrefix = GetDictString(templateCfg, "file_prefix", "")
    If Len(Trim$(filePrefix)) > 0 Then outputBase = outputBase & "_" & filePrefix
    outputPath = BuildAvailableOutputPath(outputFolder, outputBase)

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

    SaveDocumentCompat doc, outputPath

    doc.Close wdDoNotSaveChanges

    If createdWord Then wordApp.Quit wdDoNotSaveChanges

    Set doc = Nothing
    Set wordApp = Nothing

    RenderTemplate = outputPath
End Function

Private Sub SaveDocumentCompat(ByVal doc As Object, ByVal outputPath As String)
    On Error Resume Next
    CallByName doc, "SaveAs2", VbMethod, outputPath, wdFormatXMLDocument
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    Err.Clear
    CallByName doc, "SaveAs", VbMethod, outputPath, wdFormatXMLDocument
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "SaveDocumentCompat", Err.Description
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyContextToDocument(ByVal doc As Object, ByVal ctx As Object)
    Dim story As Object

    For Each story In doc.StoryRanges
        ApplyContextToStoryRange story, ctx
    Next story
End Sub

Private Sub ApplyScalarReplacements(ByVal rng As Object, ByVal ctx As Object)
    Dim key As Variant
    Dim valueText As String

    For Each key In ctx.Keys
        valueText = WordReplaceText(ctx(key))
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
    Dim findPattern As String

    leftPart = IIf(allowLeftSpaces, "[ ]@", "")
    rightPart = IIf(allowRightSpaces, "[ ]@", "")

    findPattern = "\{\{" & leftPart & EscapeWordFindText(tokenName) & rightPart & "\}\}"
    ReplaceMatchesInRange rng, findPattern, replaceText, True
End Sub

Private Sub ReplaceAllInRange(ByVal rng As Object, ByVal findText As String, ByVal replaceText As String)
    ReplaceMatchesInRange rng, findText, replaceText, False
End Sub

Private Sub ReplaceMatchesInRange(ByVal sourceRange As Object, ByVal findText As String, ByVal replaceText As String, ByVal useWildcards As Boolean)
    Dim searchRange As Object

    Set searchRange = sourceRange.Duplicate

    With searchRange.Find
        .ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = 0
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = useWildcards
    End With

    Do While searchRange.Find.Execute
        searchRange.Text = replaceText
        searchRange.Collapse wdCollapseEnd
    Loop

    Set searchRange = Nothing
End Sub

Private Sub ApplyContextToStoryRange(ByVal storyRange As Object, ByVal ctx As Object)
    Dim currentRange As Object

    Set currentRange = storyRange
    Do While Not currentRange Is Nothing
        ApplyScalarReplacements currentRange, ctx
        Set currentRange = currentRange.NextStoryRange
    Loop
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
