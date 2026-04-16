Attribute VB_Name = "Rendering"
Option Explicit

' Purpose: Render one DOCX template by substituting all {{key}} placeholders from ctx,
'          then save the result to the versioned output path.
' Inputs:  templateCfg = template row dict, ctx = INPUT context dict, wb = workbook,
'          sharedWordApp = optional already-open Word.Application (pass Nothing to let
'          RenderTemplate manage its own instance).
' Outputs: Full path of the saved output DOCX.
Public Function RenderTemplate(ByVal templateCfg As Object, ByVal ctx As Object, _
                                ByVal wb As Workbook, _
                                Optional ByVal sharedWordApp As Object = Nothing) As String
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

    ' Use a shared Word instance when provided; otherwise create/destroy locally
    If sharedWordApp Is Nothing Then
        On Error Resume Next
        Set wordApp = GetObject(, "Word.Application")
        On Error GoTo 0
        If wordApp Is Nothing Then
            Set wordApp = CreateObject("Word.Application")
            createdWord = True
        End If
        wordApp.Visible = False
    Else
        Set wordApp = sharedWordApp
        createdWord = False
    End If

    Set doc = wordApp.Documents.Open(templatePath, False, False)

    ApplyContextToDocument doc, ctx

    SaveDocumentCompat doc, outputPath

    doc.Close wdDoNotSaveChanges

    If createdWord Then wordApp.Quit wdDoNotSaveChanges

    Set doc = Nothing
    If sharedWordApp Is Nothing Then Set wordApp = Nothing

    RenderTemplate = outputPath
End Function

' Purpose: Save doc as OOXML; try SaveAs2 (Word 2013+) then fall back to SaveAs.
' Inputs:  doc = open Word document, outputPath = destination path.
' Outputs: Saves the file; raises on failure.
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

' Purpose: Walk every story range in the document and apply placeholder replacements.
' Inputs:  doc = open Word document, ctx = context dictionary.
' Outputs: Modifies document in-place.
Private Sub ApplyContextToDocument(ByVal doc As Object, ByVal ctx As Object)
    Dim story As Object

    For Each story In doc.StoryRanges
        ApplyContextToStoryRange story, ctx
    Next story
End Sub

' Purpose: Replace {{key:num}} placeholders with Vietnamese-formatted numbers.
'          Looks up each context key, parses the value as a number via ToNumber,
'          and substitutes FormatVN output (e.g. 1000000 → "1.000.000").
'          If the value cannot be parsed as a number the raw value is substituted
'          so no placeholder is left orphaned in the document.
'          Call this BEFORE ApplyScalarReplacements so that plain {{key}} tokens
'          are not consumed first.
' Inputs:  rng = Word Range, ctx = context dictionary.
' Outputs: Modifies range in-place.
Private Sub ApplyNumericReplacements(ByVal rng As Object, ByVal ctx As Object)
    ' Fast exit: skip entirely if this range has no :num placeholder markers
    If Not RangeContainsText(rng, ":num") Then Exit Sub

    Dim key As Variant
    Dim rawValue As String
    Dim formattedValue As String
    Dim numValue As Double

    For Each key In ctx.Keys
        rawValue = CellText(ctx(key))
        numValue = ToNumber(rawValue)
        ' Substitute formatted number when parseable, otherwise fall back to raw value
        ' so the placeholder is always consumed and never left in the output document.
        If numValue <> 0 Or Trim$(rawValue) = "0" Then
            formattedValue = FormatVN(numValue)
        Else
            formattedValue = WordReplaceText(ctx(key))
        End If
        ReplaceTokenInRange rng, CStr(key) & ":num", formattedValue
        If Not RangeContainsText(rng, ":num") Then Exit For
    Next key
End Sub

' Purpose: Replace all {{key}} variants in a single range using the context dictionary.
'          Two-level optimisation:
'            1. Fast-exits the entire sub when the range contains no "{{" at all.
'            2. After each key replacement, re-checks for remaining "{{" and exits early
'               once every placeholder in this range has been consumed.
'          Uses a loop-based replace (not wdReplaceAll) to support replacement text
'          longer than 255 characters.
' Inputs:  rng = Word Range, ctx = context dictionary.
' Outputs: Modifies range in-place.
Private Sub ApplyScalarReplacements(ByVal rng As Object, ByVal ctx As Object)
    ' Fast exit: skip entirely if this range has no placeholder markers
    If Not RangeContainsText(rng, "{{") Then Exit Sub

    Dim key As Variant
    Dim valueText As String

    For Each key In ctx.Keys
        valueText = WordReplaceText(ctx(key))
        ReplaceTokenInRange rng, CStr(key), valueText
        ' Early exit: stop once all placeholders in this range are consumed.
        ' Critical for large context dictionaries (20+ templates, 50+ keys):
        ' avoids executing 7 Find operations per remaining key after the last
        ' placeholder has already been replaced.
        If Not RangeContainsText(rng, "{{") Then Exit For
    Next key
End Sub

' Purpose: Test whether a Word Range contains a literal text string.
' Inputs:  rng = range to search, searchText = literal string to find.
' Outputs: True if the text was found anywhere in the range.
Private Function RangeContainsText(ByVal rng As Object, ByVal searchText As String) As Boolean
    Dim checkRange As Object

    Set checkRange = rng.Duplicate
    With checkRange.Find
        .ClearFormatting
        .Text = searchText
        .Forward = True
        .Wrap = 0
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
    End With
    RangeContainsText = checkRange.Find.Execute
    Set checkRange = Nothing
End Function

' Purpose: Replace all spacing variants of {{tokenName}} with replaceText.
'          Exact-match passes are run first; they handle replacement text of any
'          length (Word wdReplaceAll truncates text beyond 255 characters).
'          Wildcard passes are run after to catch remaining spacing variants.
' Inputs:  rng = Word Range, tokenName = placeholder key, replaceText = substitution value.
' Outputs: Modifies range in-place.
Private Sub ReplaceTokenInRange(ByVal rng As Object, ByVal tokenName As String, ByVal replaceText As String)
    ' Exact-match passes (also safe for >255-char replacement text)
    ReplaceAllInRange rng, "{{" & tokenName & "}}", replaceText
    ReplaceAllInRange rng, "{{ " & tokenName & " }}", replaceText
    ReplaceAllInRange rng, "{{" & tokenName & " }}", replaceText
    ReplaceAllInRange rng, "{{ " & tokenName & "}}", replaceText

    ' Wildcard passes catch any remaining spacing variants
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

' Purpose: Find-and-replace using an explicit loop so replacement text of any length is
'          supported (Word's built-in wdReplaceAll truncates text beyond 255 characters).
' Inputs:  sourceRange = story range to search, findText = search term or pattern,
'          replaceText = replacement string (unlimited length),
'          useWildcards = whether findText is a Word wildcard pattern.
' Outputs: Modifies sourceRange in-place.
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

' Purpose: Traverse all linked story ranges (e.g. multiple header sections) and apply replacements.
' Inputs:  storyRange = first range of a story type, ctx = context dictionary.
' Outputs: Modifies all linked story ranges in-place.
Private Sub ApplyContextToStoryRange(ByVal storyRange As Object, ByVal ctx As Object)
    Dim currentRange As Object

    Set currentRange = storyRange
    Do While Not currentRange Is Nothing
        ApplyNumericReplacements currentRange, ctx
        ApplyScalarReplacements currentRange, ctx
        Set currentRange = currentRange.NextStoryRange
    Loop
End Sub

' Purpose: Find a free versioned output path (never overwrites existing files).
' Inputs:  outputRoot = output folder path, filePrefix = base name without extension.
' Outputs: Full path string ending in _v<N>.docx where N is the lowest free number.
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
