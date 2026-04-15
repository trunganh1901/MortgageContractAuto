Attribute VB_Name = "modHistory"
Option Explicit

' =============================================================================
' modHistory — Load a historical export run back into the INPUT sheet.
'
' Entry points (call from a button or Immediate window):
'   LoadFromHistory      — file-picker → JSON → restore INPUT column C
'   BrowseExportHistory  — pick from a numbered list of recent runs (CSV-based)
' =============================================================================

' Purpose: Show a file-picker dialog so the user can select any JSON log file,
'          parse its "context" snapshot, and restore every matching key in the
'          INPUT sheet's column C.
' Inputs:  none (reads wb = ThisWorkbook, INPUT sheet)
' Outputs: Overwrites INPUT column C values for keys found in the log.
Public Sub LoadFromHistory()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim filePath As String
    Dim jsonText As String
    Dim runId As String
    Dim restoredCount As Long

    Set wb = ThisWorkbook
    filePath = PickLogFile(wb)
    If Len(filePath) = 0 Then Exit Sub  ' user cancelled

    If MsgBox( _
        "This will overwrite editable fields (column C) on the INPUT sheet" & vbCrLf & _
        "with values from the selected log file." & vbCrLf & vbCrLf & _
        "File: " & vbCrLf & filePath & vbCrLf & vbCrLf & _
        "Continue?", _
        vbQuestion + vbYesNo, "Load Historical Data") = vbNo Then
        Exit Sub
    End If

    jsonText = ReadTextFileUtf8(filePath)
    If Len(Trim$(jsonText)) = 0 Then
        MsgBox "The selected file is empty or could not be read.", vbExclamation, "Load Historical Data"
        Exit Sub
    End If

    runId = ExtractJsonString(jsonText, "run_id")
    restoredCount = RestoreContextFromJson(jsonText, wb)

    MsgBox "Restored " & restoredCount & " field(s) from run:" & vbCrLf & runId, _
        vbInformation, "Load Historical Data"

    Exit Sub

ErrorHandler:
    MsgBox "Load failed: " & Err.Description, vbCritical, "LoadFromHistory"
End Sub

' Purpose: Read the export_history.csv from the Logs/<YYYY> folders, present
'          the most recent runs to the user as a numbered list, and load the
'          selected run's JSON file back into the INPUT sheet.
' Inputs:  none (reads wb = ThisWorkbook)
' Outputs: Overwrites INPUT column C values for the chosen run.
Public Sub BrowseExportHistory()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim logsRoot As String
    Dim runs() As String          ' each element: "run_id|json_path|started_at|status|customer|cif"
    Dim runCount As Long
    Dim listText As String
    Dim choice As String
    Dim choiceNum As Long
    Dim jsonPath As String
    Dim jsonText As String
    Dim restoredCount As Long

    Set wb = ThisWorkbook
    logsRoot = BuildPath(wb.Path, "Logs")

    If Dir$(logsRoot, vbDirectory) = vbNullString Then
        MsgBox "No Logs folder found at:" & vbCrLf & logsRoot, vbExclamation, "Export History"
        Exit Sub
    End If

    CollectRecentRuns logsRoot, runs, runCount

    If runCount = 0 Then
        MsgBox "No export history found in:" & vbCrLf & logsRoot, vbInformation, "Export History"
        Exit Sub
    End If

    listText = "Recent exports (newest first). Enter the number to load:" & vbCrLf & vbCrLf
    Dim idx As Long
    For idx = 1 To runCount
        Dim parts() As String
        parts = Split(runs(idx - 1), "|")
        listText = listText & idx & ".  " & _
            parts(2) & "  " & _
            "[" & parts(3) & "]  " & _
            "CIF=" & parts(5) & "  " & parts(4) & vbCrLf
    Next idx

    choice = InputBox(listText, "Browse Export History", "1")
    If Len(Trim$(choice)) = 0 Then Exit Sub

    If Not IsNumeric(choice) Then
        MsgBox "Please enter a number.", vbExclamation, "Browse Export History"
        Exit Sub
    End If

    choiceNum = CLng(choice)
    If choiceNum < 1 Or choiceNum > runCount Then
        MsgBox "Number out of range (1–" & runCount & ").", vbExclamation, "Browse Export History"
        Exit Sub
    End If

    parts = Split(runs(choiceNum - 1), "|")
    jsonPath = parts(1)

    If Dir$(jsonPath, vbNormal) = vbNullString Then
        MsgBox "JSON file not found:" & vbCrLf & jsonPath, vbExclamation, "Browse Export History"
        Exit Sub
    End If

    If MsgBox( _
        "Load run from " & parts(2) & "?" & vbCrLf & _
        "CIF: " & parts(5) & "   Customer: " & parts(4) & vbCrLf & vbCrLf & _
        "This will overwrite editable fields (column C) on the INPUT sheet." & vbCrLf & vbCrLf & _
        "Continue?", _
        vbQuestion + vbYesNo, "Load Historical Data") = vbNo Then
        Exit Sub
    End If

    jsonText = ReadTextFileUtf8(jsonPath)
    Dim runId As String
    runId = ExtractJsonString(jsonText, "run_id")
    restoredCount = RestoreContextFromJson(jsonText, wb)

    MsgBox "Restored " & restoredCount & " field(s) from run:" & vbCrLf & runId, _
        vbInformation, "Load Historical Data"

    Exit Sub

ErrorHandler:
    MsgBox "Browse failed: " & Err.Description, vbCritical, "BrowseExportHistory"
End Sub

' =============================================================================
' Private helpers — file picking and context restoration
' =============================================================================

' Purpose: Show the OS file-picker dialog starting in the Logs folder.
' Inputs:  wb = current workbook.
' Outputs: Full path of selected file, or empty string if cancelled.
Private Function PickLogFile(ByVal wb As Workbook) As String
    Dim dlg As FileDialog
    Dim logsRoot As String

    logsRoot = BuildPath(wb.Path, "Logs")

    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Select a JSON log file"
        .Filters.Clear
        .Filters.Add "JSON log files", "*.json"
        .AllowMultiSelect = False
        If Dir$(logsRoot, vbDirectory) <> vbNullString Then
            .InitialFileName = logsRoot & "\"
        End If
        If .Show = True Then PickLogFile = .SelectedItems(1)
    End With
End Function

' Purpose: Parse the "context" block from a JSON log file and write values
'          back into INPUT sheet column C for every key that exists there.
' Inputs:  jsonText = full file content, wb = workbook with INPUT sheet.
' Outputs: Number of INPUT rows updated.
Private Function RestoreContextFromJson(ByVal jsonText As String, ByVal wb As Workbook) As Long
    Dim ctxJson As String
    Dim ctx As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNo As Long
    Dim keyText As String
    Dim restoredCount As Long

    ctxJson = ExtractJsonObjectBlock(jsonText, "context")
    If Len(ctxJson) = 0 Then
        MsgBox "No context data found in the log file.", vbExclamation, "Load Historical Data"
        Exit Function
    End If

    Set ctx = ParseFlatJsonObject(ctxJson)

    Set ws = wb.Sheets("INPUT")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    For rowNo = 1 To lastRow
        keyText = Trim$(CellText(ws.Cells(rowNo, "A").Value))
        If Len(keyText) > 0 And ctx.Exists(keyText) Then
            ws.Cells(rowNo, "C").Value = ctx(keyText)
            restoredCount = restoredCount + 1
        End If
    Next rowNo

    Application.ScreenUpdating = True

    RestoreContextFromJson = restoredCount
End Function

' Purpose: Walk Logs/<YYYY>/export_history.csv files and return up to 20
'          most recent run rows as pipe-delimited strings.
' Inputs:  logsRoot = path to the Logs folder, runs() = output array (resized
'          here), runCount = number of valid rows written.
' Outputs: Populates runs() and runCount.
Private Sub CollectRecentRuns(ByVal logsRoot As String, ByRef runs() As String, ByRef runCount As Long)
    Const MAX_RUNS As Long = 20

    Dim fso As Object
    Dim yearFolders As Object
    Dim yearFolder As Object
    Dim csvPath As String
    Dim csvLines() As String
    Dim lineNo As Long
    Dim cols() As String

    ReDim runs(0 To MAX_RUNS - 1)
    runCount = 0

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(logsRoot) Then Exit Sub

    Set yearFolders = fso.GetFolder(logsRoot).SubFolders

    ' Iterate year folders newest-first
    Dim yearNames() As String
    Dim yearCount As Long
    yearCount = 0
    ReDim yearNames(yearFolders.Count - 1)

    Dim yf As Object
    For Each yf In yearFolders
        yearNames(yearCount) = yf.Name
        yearCount = yearCount + 1
    Next yf

    ' Sort descending (simple bubble sort — only a handful of year folders)
    Dim tmp As String
    Dim a As Long, b As Long
    For a = 0 To yearCount - 2
        For b = a + 1 To yearCount - 1
            If yearNames(a) < yearNames(b) Then
                tmp = yearNames(a) : yearNames(a) = yearNames(b) : yearNames(b) = tmp
            End If
        Next b
    Next a

    For a = 0 To yearCount - 1
        csvPath = BuildPath(logsRoot, yearNames(a), "export_history.csv")
        If Dir$(csvPath, vbNormal) = vbNullString Then GoTo NextYear

        csvLines = Split(ReadTextFileUtf8(csvPath), vbLf)

        ' Read from the bottom (newest rows) upward; skip header (line 0)
        For lineNo = UBound(csvLines) To 1 Step -1
            Dim rawLine As String
            rawLine = Trim$(csvLines(lineNo))
            If Len(rawLine) = 0 Then GoTo NextLine

            cols = ParseCsvLine(rawLine)
            ' CSV columns (0-indexed):
            ' 0=run_id, 1=status, 2=started_at, 3=finished_at, 4=duration_seconds,
            ' 5=template_code, 6=template_description, 7=template_file, 8=file_prefix,
            ' 9=output_name, 10=output_path, 11=customer_name, 12=cif, 13=stt_hd,
            ' 14=context_key_count, 15=error_message, 16=json_path
            If UBound(cols) < 16 Then GoTo NextLine

            runs(runCount) = cols(0) & "|" & cols(16) & "|" & cols(2) & "|" & cols(1) & "|" & cols(11) & "|" & cols(12)
            runCount = runCount + 1
            If runCount >= MAX_RUNS Then Exit Sub

NextLine:
        Next lineNo

NextYear:
    Next a
End Sub

' =============================================================================
' Private helpers — JSON parsing
' =============================================================================

' Purpose: Extract the raw JSON object block for a named top-level key,
'          e.g. the value of "context": {...} from the full log JSON.
' Inputs:  jsonText = full JSON string, key = property name.
' Outputs: Raw JSON object string including braces, or empty string.
Private Function ExtractJsonObjectBlock(ByVal jsonText As String, ByVal key As String) As String
    Dim searchFor As String
    Dim keyPos As Long
    Dim braceStart As Long
    Dim depth As Long
    Dim i As Long
    Dim ch As String
    Dim inString As Boolean
    Dim prevCh As String

    searchFor = """" & key & """:"
    keyPos = InStr(1, jsonText, searchFor, vbBinaryCompare)
    If keyPos = 0 Then Exit Function

    braceStart = InStr(keyPos + Len(searchFor), jsonText, "{")
    If braceStart = 0 Then Exit Function

    depth = 0
    inString = False
    prevCh = vbNullString

    For i = braceStart To Len(jsonText)
        ch = Mid$(jsonText, i, 1)

        If inString Then
            If ch = """" And prevCh <> "\" Then inString = False
        Else
            Select Case ch
                Case """" : inString = True
                Case "{" : depth = depth + 1
                Case "}"
                    depth = depth - 1
                    If depth = 0 Then
                        ExtractJsonObjectBlock = Mid$(jsonText, braceStart, i - braceStart + 1)
                        Exit Function
                    End If
            End Select
        End If

        prevCh = ch
    Next i
End Function

' Purpose: Parse a flat JSON object  {"key":"value", "key2":123, ...}
'          into a case-insensitive Scripting.Dictionary of string values.
'          Handles string escapes; numeric/boolean/null values stored as text.
' Inputs:  jsonText = JSON object string (must start with '{').
' Outputs: Dictionary with string keys and string values.
Private Function ParseFlatJsonObject(ByVal jsonText As String) As Object
    Const ST_KEY As Integer = 0
    Const ST_COLON As Integer = 1
    Const ST_VALUE As Integer = 2
    Const ST_COMMA As Integer = 3

    Dim result As Object
    Dim i As Long
    Dim ch As String
    Dim state As Integer
    Dim currentKey As String
    Dim strVal As String
    Dim bareVal As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    i = InStr(jsonText, "{")
    If i = 0 Then Set ParseFlatJsonObject = result : Exit Function
    i = i + 1  ' skip opening brace

    state = ST_KEY
    currentKey = vbNullString

    Do While i <= Len(jsonText)
        ch = Mid$(jsonText, i, 1)

        Select Case ch
            Case " ", vbTab, vbCr, vbLf
                i = i + 1

            Case "}"
                Exit Do

            Case ","
                state = ST_KEY
                currentKey = vbNullString
                i = i + 1

            Case ":"
                state = ST_VALUE
                i = i + 1

            Case """"
                strVal = ReadJsonString(jsonText, i)  ' i advanced past closing quote
                If state = ST_KEY Then
                    currentKey = strVal
                    state = ST_COLON
                ElseIf state = ST_VALUE Then
                    If Len(currentKey) > 0 Then result(currentKey) = strVal
                    state = ST_COMMA
                End If

            Case Else
                If state = ST_VALUE Then
                    bareVal = ReadJsonBareValue(jsonText, i)  ' i advanced
                    If Len(currentKey) > 0 Then result(currentKey) = bareVal
                    state = ST_COMMA
                Else
                    i = i + 1
                End If
        End Select
    Loop

    Set ParseFlatJsonObject = result
End Function

' Purpose: Read a JSON string literal starting at the opening quote at pos.
'          Handles escape sequences: \", \\, \/, \n, \r, \t, \uXXXX.
'          Advances pos to the character after the closing quote.
' Inputs:  jsonText = full JSON string, pos = index of opening quote (ByRef).
' Outputs: Unescaped string value.
Private Function ReadJsonString(ByVal jsonText As String, ByRef pos As Long) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim escaped As Boolean

    i = pos + 1  ' skip opening quote
    escaped = False

    Do While i <= Len(jsonText)
        ch = Mid$(jsonText, i, 1)

        If escaped Then
            Select Case ch
                Case """" : result = result & """"
                Case "\"  : result = result & "\"
                Case "/"  : result = result & "/"
                Case "n"  : result = result & vbLf
                Case "r"  : result = result & vbCr
                Case "t"  : result = result & vbTab
                Case "u"
                    If i + 4 <= Len(jsonText) Then
                        result = result & ChrW$(CLng("&H" & Mid$(jsonText, i + 1, 4)))
                        i = i + 4
                    End If
                Case Else : result = result & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            pos = i + 1
            ReadJsonString = result
            Exit Function
        Else
            result = result & ch
        End If

        i = i + 1
    Loop

    pos = i
    ReadJsonString = result
End Function

' Purpose: Read a bare JSON value (number, true, false, null) starting at pos.
'          Stops at delimiters: , } ] or whitespace. Advances pos past the value.
' Inputs:  jsonText = full JSON string, pos = start index (ByRef).
' Outputs: Raw token string (e.g. "42", "true", "null").
Private Function ReadJsonBareValue(ByVal jsonText As String, ByRef pos As Long) As String
    Dim result As String
    Dim ch As String

    Do While pos <= Len(jsonText)
        ch = Mid$(jsonText, pos, 1)
        Select Case ch
            Case ",", "}", "]", " ", vbTab, vbCr, vbLf
                Exit Do
            Case Else
                result = result & ch
                pos = pos + 1
        End Select
    Loop

    ReadJsonBareValue = result
End Function

' Purpose: Extract a top-level scalar string value for a named key.
'          Useful for pulling metadata fields like "run_id" or "started_at".
' Inputs:  jsonText = full JSON string, key = property name.
' Outputs: Unescaped string value, or empty string if not found.
Private Function ExtractJsonString(ByVal jsonText As String, ByVal key As String) As String
    Dim searchFor As String
    Dim keyPos As Long
    Dim quotePos As Long
    Dim dummyPos As Long

    searchFor = """" & key & """:"
    keyPos = InStr(1, jsonText, searchFor, vbBinaryCompare)
    If keyPos = 0 Then Exit Function

    quotePos = InStr(keyPos + Len(searchFor), jsonText, """")
    If quotePos = 0 Then Exit Function

    dummyPos = quotePos
    ExtractJsonString = ReadJsonString(jsonText, dummyPos)
End Function

' =============================================================================
' Private helpers — CSV parsing
' =============================================================================

' Purpose: Parse one CSV line into an array of unquoted field strings.
'          Handles RFC-4180 double-quote escaping ("" inside a quoted field).
' Inputs:  line = one CSV text line (no trailing CRLF).
' Outputs: 0-based string array of field values.
Private Function ParseCsvLine(ByVal line As String) As String()
    Dim fields() As String
    Dim fieldCount As Long
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim current As String

    ReDim fields(0 To 50)
    fieldCount = 0
    inQuotes = False
    current = vbNullString
    i = 1

    Do While i <= Len(line)
        ch = Mid$(line, i, 1)

        If inQuotes Then
            If ch = """" Then
                If i + 1 <= Len(line) And Mid$(line, i + 1, 1) = """" Then
                    current = current & """"  ' escaped quote
                    i = i + 1
                Else
                    inQuotes = False  ' closing quote
                End If
            Else
                current = current & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQuotes = True
                Case ","
                    fields(fieldCount) = current
                    fieldCount = fieldCount + 1
                    If fieldCount > UBound(fields) Then ReDim Preserve fields(0 To fieldCount + 10)
                    current = vbNullString
                Case Else
                    current = current & ch
            End Select
        End If

        i = i + 1
    Loop

    fields(fieldCount) = current
    ReDim Preserve fields(0 To fieldCount)
    ParseCsvLine = fields
End Function

' =============================================================================
' Private helpers — file I/O (local copies to keep module self-contained)
' =============================================================================

' Purpose: Read a UTF-8 encoded text file via ADODB.Stream.
' Inputs:  filePath = full path to the file.
' Outputs: File contents as a VBA string.
Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadTextFileUtf8 = stream.ReadText
    stream.Close
End Function
