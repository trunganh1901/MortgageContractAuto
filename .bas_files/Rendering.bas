Attribute VB_Name = "Rendering"
Option Explicit

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
            ' Keep late-bound collection access stable during import/compile.
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
    Dim rowNo As Long
    Dim rowText As String
    Dim currentRow As Long
    Dim itemIndex As Long

    startRow = 0
    endRow = 0

    For rowNo = 1 To tbl.Rows.Count
        rowText = CleanWordCellText(tbl.Rows(rowNo).Range.Text)
        If InStr(1, rowText, "for item in items", vbTextCompare) > 0 Then startRow = rowNo
        If InStr(1, rowText, "endfor", vbTextCompare) > 0 Then
            endRow = rowNo
            Exit For
        End If
    Next rowNo

    If startRow = 0 Or endRow = 0 Then Exit Sub
    If endRow - startRow <> 2 Then Exit Sub

    currentRow = startRow + 1

    If items.Count = 0 Then
        tbl.Rows(endRow).Delete
        tbl.Rows(currentRow).Delete
        tbl.Rows(startRow).Delete
        Exit Sub
    End If

    For rowNo = 2 To items.Count
        tbl.Rows(currentRow).Select
        tbl.Application.Selection.InsertRowsBelow 1
    Next rowNo

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
