Attribute VB_Name = "HELPER"
Option Explicit

' Purpose: Write a progress message to the Excel status bar so users can see what is running.
'          Calls DoEvents so the status bar repaints immediately.
' Inputs:  message = text to display.
' Outputs: Updates Application.StatusBar.
Public Sub ShowStatus(ByVal message As String)
    Application.StatusBar = message
    DoEvents
End Sub

' Purpose: Restore the default Excel status bar (clears any message set by ShowStatus).
' Inputs:  None.
' Outputs: Resets Application.StatusBar to False (Excel default).
Public Sub ClearStatus()
    Application.StatusBar = False
End Sub

' Purpose: After a successful export, ask the user whether to open the output folder.
' Inputs:  outputPath = full path of the generated DOCX (optional; defaults to Output\ root).
' Outputs: Opens Explorer if the user chooses Yes.
Public Sub PromptOpenOutputFolder(Optional ByVal outputPath As String = "")
    Dim folderPath As String
    Dim userChoice As VbMsgBoxResult

    If Len(Trim$(outputPath)) > 0 Then
        folderPath = Left$(outputPath, InStrRev(outputPath, "\") - 1)
    Else
        folderPath = BuildPath(ThisWorkbook.Path, "Output")
    End If

    If Len(folderPath) = 0 Then
        MsgBox "No output folder path found.", vbExclamation
        Exit Sub
    End If

    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder not found:" & vbCrLf & folderPath, vbExclamation
        Exit Sub
    End If

    userChoice = MsgBox( _
        "Export completed successfully." & vbCrLf & vbCrLf & _
        "Output folder:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
        "Open this folder now?", _
        vbQuestion + vbYesNo, _
        "Done" _
    )

    If userChoice = vbYes Then
        If Len(folderPath) > 0 And Dir(folderPath, vbDirectory) <> "" Then
            Shell "explorer.exe """ & folderPath & """", vbNormalFocus
        Else
            MsgBox "Output folder not found.", vbExclamation
        End If
    End If
End Sub

Public Sub OpenCustomerOutputFolder()
    PromptOpenOutputFolder
End Sub

Function IncrementSTT_HD() As String
    Dim ws As Worksheet
    Dim r As Long
    Dim currentVal As Long

    Set ws = ThisWorkbook.ActiveSheet

    For r = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If Trim(UCase(ws.Cells(r, "A").Value)) = "STT_HD" Then

            If IsNumeric(ws.Cells(r, "C").Value) Then
                currentVal = CLng(ws.Cells(r, "C").Value)
            Else
                currentVal = 0
            End If

            ws.Cells(r, "C").Value = Format(currentVal + 1, "00")
            IncrementSTT_HD = ws.Cells(r, "C").Value

            Exit Function
        End If
    Next r
End Function

Sub DisableAllButtons(Optional ws As Worksheet)
    Dim shp As Shape
    If ws Is Nothing Then Set ws = ActiveSheet

    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            shp.ControlFormat.Enabled = False
        End If
    Next shp
End Sub

Sub EnableAllButtons(Optional ws As Worksheet)
    Dim shp As Shape
    If ws Is Nothing Then Set ws = ActiveSheet

    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            shp.ControlFormat.Enabled = True
        End If
    Next shp
End Sub
