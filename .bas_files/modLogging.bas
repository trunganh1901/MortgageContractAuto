Attribute VB_Name = "modLogging"
Option Explicit

' Purpose: Create one export-run log payload.
' Inputs: wb = current workbook, ctx = current INPUT context, startedAt = run start time.
' Outputs: Dictionary with run metadata and a cloned context snapshot.
Public Function CreateExportLog(ByVal wb As Workbook, ByVal ctx As Object, ByVal startedAt As Date) As Object
    Dim logEntry As Object
    Dim contextSnapshot As Object
    Dim logRoot As String

    Set logEntry = CreateObject("Scripting.Dictionary")
    logEntry.CompareMode = 1
    Set contextSnapshot = CloneDictionary(ctx)
    logRoot = BuildPath(wb.Path, "Logs")

    logEntry("run_id") = BuildRunId(startedAt)
    logEntry("status") = "running"
    logEntry("started_at") = IsoTimestamp(startedAt)
    logEntry("started_serial") = CDbl(startedAt)
    logEntry("finished_at") = vbNullString
    logEntry("duration_seconds") = 0
    logEntry("workbook_name") = wb.Name
    logEntry("workbook_path") = wb.FullName
    logEntry("log_folder") = BuildStructuredFolder(logRoot, ctx, "document")
    logEntry.Add "context", contextSnapshot
    logEntry("error_message") = vbNullString

    Set CreateExportLog = logEntry
End Function

' Purpose: Build one output record for the current run.
' Inputs: templateCfg = selected template configuration, outputPath = generated DOCX path.
' Outputs: Dictionary with template and output file metadata.
Public Function CreateExportOutput(ByVal templateCfg As Object, ByVal outputPath As String) As Object
    Dim outputEntry As Object

    Set outputEntry = CreateObject("Scripting.Dictionary")
    outputEntry.CompareMode = 1
    outputEntry("template_code") = GetDictString(templateCfg, "template_code")
    outputEntry("template_description") = GetDictString(templateCfg, "description")
    outputEntry("template_file") = GetDictString(templateCfg, "docx_file")
    outputEntry("file_prefix") = GetDictString(templateCfg, "file_prefix")
    outputEntry("output_path") = outputPath
    outputEntry("output_name") = FileNameFromPath(outputPath)

    Set CreateExportOutput = outputEntry
End Function

' Purpose: Finalize the run state and write CSV + JSON log artifacts.
' Inputs: logEntry = run metadata, outputs = collection of output dictionaries, finishedAt = run end time, status = success/failed, errorMessage = optional failure text.
' Outputs: Writes files to Logs folder. No return value.
Public Sub SaveExportLog(ByVal logEntry As Object, ByVal outputs As Collection, ByVal finishedAt As Date, ByVal status As String, Optional ByVal errorMessage As String = "")
    Dim logFolder As String
    Dim jsonPath As String
    Dim csvPath As String

    logEntry("status") = status
    logEntry("finished_at") = IsoTimestamp(finishedAt)
    logEntry("duration_seconds") = DateDiff("s", CDate(logEntry("started_serial")), finishedAt)
    logEntry("error_message") = errorMessage

    logFolder = GetDictString(logEntry, "log_folder")
    EnsureFolderTreeExists logFolder

    jsonPath = BuildPath(logFolder, GetDictString(logEntry, "run_id") & ".json")
    csvPath = BuildPath(logFolder, "export_history.csv")

    WriteTextFileUtf8 jsonPath, ExportLogToJson(logEntry, outputs)
    AppendExportHistoryCsv csvPath, logEntry, outputs, jsonPath
End Sub

' Purpose: Append one or more rows to the export history CSV.
' Inputs: csvPath = destination CSV file, logEntry = run metadata, outputs = output records, jsonPath = run JSON file path.
' Outputs: Updates the CSV file in UTF-8 format.
Private Sub AppendExportHistoryCsv(ByVal csvPath As String, ByVal logEntry As Object, ByVal outputs As Collection, ByVal jsonPath As String)
    Dim lines As Collection
    Dim rowItem As Variant
    Dim hasFile As Boolean

    Set lines = New Collection
    hasFile = (Dir$(csvPath, vbNormal) <> vbNullString)

    If Not hasFile Then
        lines.Add CsvRow(Array( _
            "run_id", "status", "started_at", "finished_at", "duration_seconds", _
            "template_code", "template_description", "template_file", "file_prefix", _
            "output_name", "output_path", "customer_name", "cif", "stt_hd", _
            "context_key_count", "error_message", "json_path"))
    End If

    If outputs.Count = 0 Then
        lines.Add CsvRow(Array( _
            GetDictString(logEntry, "run_id"), GetDictString(logEntry, "status"), GetDictString(logEntry, "started_at"), GetDictString(logEntry, "finished_at"), CStr(logEntry("duration_seconds")), _
            vbNullString, vbNullString, vbNullString, vbNullString, _
            vbNullString, vbNullString, ContextValue(logEntry, "customer_name"), ContextValue(logEntry, "CIF"), ContextValue(logEntry, "stt_hd"), _
            CStr(DictCount(logEntry("context"))), GetDictString(logEntry, "error_message"), jsonPath))
    Else
        For Each rowItem In outputs
            lines.Add CsvRow(Array( _
                GetDictString(logEntry, "run_id"), GetDictString(logEntry, "status"), GetDictString(logEntry, "started_at"), GetDictString(logEntry, "finished_at"), CStr(logEntry("duration_seconds")), _
                GetDictString(rowItem, "template_code"), GetDictString(rowItem, "template_description"), GetDictString(rowItem, "template_file"), GetDictString(rowItem, "file_prefix"), _
                GetDictString(rowItem, "output_name"), GetDictString(rowItem, "output_path"), ContextValue(logEntry, "customer_name"), ContextValue(logEntry, "CIF"), ContextValue(logEntry, "stt_hd"), _
                CStr(DictCount(logEntry("context"))), GetDictString(logEntry, "error_message"), jsonPath))
        Next rowItem
    End If

    AppendTextFileUtf8 csvPath, JoinCollection(lines, vbCrLf) & vbCrLf
End Sub

' Purpose: Serialize one run into JSON for machine parsing.
' Inputs: logEntry = run metadata, outputs = output records.
' Outputs: JSON string.
Private Function ExportLogToJson(ByVal logEntry As Object, ByVal outputs As Collection) As String
    Dim parts As Collection
    Dim outputParts As Collection
    Dim item As Variant

    Set parts = New Collection
    parts.Add JsonPair("run_id", GetDictString(logEntry, "run_id"))
    parts.Add JsonPair("status", GetDictString(logEntry, "status"))
    parts.Add JsonPair("started_at", GetDictString(logEntry, "started_at"))
    parts.Add JsonPair("finished_at", GetDictString(logEntry, "finished_at"))
    parts.Add JsonPair("duration_seconds", CStr(logEntry("duration_seconds")), False)
    parts.Add JsonPair("workbook_name", GetDictString(logEntry, "workbook_name"))
    parts.Add JsonPair("workbook_path", GetDictString(logEntry, "workbook_path"))
    parts.Add JsonPair("error_message", GetDictString(logEntry, "error_message"))
    parts.Add """context"": " & DictToJson(logEntry("context"))

    Set outputParts = New Collection
    For Each item In outputs
        outputParts.Add DictToJson(item)
    Next item
    parts.Add """outputs"": [" & JoinCollection(outputParts, ",") & "]"

    ExportLogToJson = "{" & JoinCollection(parts, ",") & "}"
End Function

' Purpose: Convert a flat dictionary to JSON.
' Inputs: dict = dictionary with scalar values.
' Outputs: JSON object string.
Private Function DictToJson(ByVal dict As Object) As String
    Dim parts As Collection
    Dim key As Variant

    Set parts = New Collection
    For Each key In dict.Keys
        parts.Add JsonPair(CStr(key), CellText(dict(key)))
    Next key

    DictToJson = "{" & JoinCollection(parts, ",") & "}"
End Function

' Purpose: Format one JSON key/value pair.
' Inputs: key = property name, value = property value, quoteValue = whether the value is a JSON string.
' Outputs: JSON fragment.
Private Function JsonPair(ByVal key As String, ByVal value As String, Optional ByVal quoteValue As Boolean = True) As String
    If quoteValue Then
        JsonPair = """" & JsonEscape(key) & """: """ & JsonEscape(value) & """"
    Else
        JsonPair = """" & JsonEscape(key) & """: " & value
    End If
End Function

' Purpose: Escape string content for JSON output.
' Inputs: value = raw text.
' Outputs: Safe JSON string content without surrounding quotes.
Private Function JsonEscape(ByVal value As String) As String
    Dim normalized As String

    normalized = value
    normalized = Replace$(normalized, "\", "\\")
    normalized = Replace$(normalized, """", Chr$(92) & """")
    normalized = Replace$(normalized, vbCrLf, Chr$(92) & "n")
    normalized = Replace$(normalized, vbCr, Chr$(92) & "n")
    normalized = Replace$(normalized, vbLf, Chr$(92) & "n")
    normalized = Replace$(normalized, vbTab, Chr$(92) & "t")

    JsonEscape = normalized
End Function

' Purpose: Convert an array of values into one CSV line.
' Inputs: values = array of scalar values.
' Outputs: One CSV row string.
Private Function CsvRow(ByVal values As Variant) As String
    Dim items() As String
    Dim i As Long

    ReDim items(LBound(values) To UBound(values))
    For i = LBound(values) To UBound(values)
        items(i) = CsvValue(CStr(values(i)))
    Next i

    CsvRow = Join(items, ",")
End Function

' Purpose: Escape one scalar value for CSV output.
' Inputs: value = raw text.
' Outputs: Quoted CSV-safe text.
Private Function CsvValue(ByVal value As String) As String
    Dim escaped As String

    escaped = Replace$(value, """", """""")
    CsvValue = """" & escaped & """"
End Function

' Purpose: Join a collection of strings with a delimiter.
' Inputs: items = collection of strings, delimiter = separator string.
' Outputs: Joined string.
Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim result As String
    Dim item As Variant

    For Each item In items
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(item)
    Next item

    JoinCollection = result
End Function

' Purpose: Clone the current INPUT context so the JSON log preserves the run state.
' Inputs: sourceDict = source dictionary.
' Outputs: New dictionary with copied scalar values.
Private Function CloneDictionary(ByVal sourceDict As Object) As Object
    Dim result As Object
    Dim key As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    For Each key In sourceDict.Keys
        result(CStr(key)) = sourceDict(key)
    Next key

    Set CloneDictionary = result
End Function

' Purpose: Read one context value from the run log safely.
' Inputs: logEntry = run metadata, key = context key.
' Outputs: Context value or empty string.
Private Function ContextValue(ByVal logEntry As Object, ByVal key As String) As String
    ContextValue = GetDictString(logEntry("context"), key)
End Function

' Purpose: Return the number of keys in a dictionary.
' Inputs: dict = dictionary object.
' Outputs: Key count.
Private Function DictCount(ByVal dict As Object) As Long
    DictCount = dict.Count
End Function

' Purpose: Build a unique file-friendly run identifier.
' Inputs: stamp = run start time.
' Outputs: Run ID string.
Private Function BuildRunId(ByVal stamp As Date) As String
    BuildRunId = "run_" & Format$(stamp, "yyyymmdd_hhnnss") & "_" & Right$("000" & CStr(CLng((Timer - Int(Timer)) * 1000)), 3)
End Function

' Purpose: Format a timestamp for logs.
' Inputs: stamp = date/time value.
' Outputs: ISO-like timestamp string.
Private Function IsoTimestamp(ByVal stamp As Date) As String
    IsoTimestamp = Format$(stamp, "yyyy-mm-dd\THH:nn:ss")
End Function

' Purpose: Extract a file name from a full path.
' Inputs: fullPath = full path string.
' Outputs: File name only.
Private Function FileNameFromPath(ByVal fullPath As String) As String
    If InStrRev(fullPath, "\") > 0 Then
        FileNameFromPath = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    Else
        FileNameFromPath = fullPath
    End If
End Function

' Purpose: Write a UTF-8 text file.
' Inputs: filePath = destination path, textContent = file contents.
' Outputs: Overwrites the destination file.
Private Sub WriteTextFileUtf8(ByVal filePath As String, ByVal textContent As String)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textContent
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

' Purpose: Append UTF-8 text to a file.
' Inputs: filePath = destination path, textContent = text to append.
' Outputs: Updates or creates the file.
Private Sub AppendTextFileUtf8(ByVal filePath As String, ByVal textContent As String)
    Dim currentText As String

    If Dir$(filePath, vbNormal) <> vbNullString Then
        currentText = ReadTextFileUtf8(filePath)
    End If

    WriteTextFileUtf8 filePath, currentText & textContent
End Sub

' Purpose: Read a UTF-8 text file.
' Inputs: filePath = source path.
' Outputs: File contents as text.
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
