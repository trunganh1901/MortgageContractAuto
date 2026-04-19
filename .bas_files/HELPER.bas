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

    For r = 1 To ws.Cells(ws.Rows.count, "A").End(xlUp).Row
        If Trim(UCase(ws.Cells(r, "A").value)) = "STT_HD" Then

            If IsNumeric(ws.Cells(r, "C").value) Then
                currentVal = CLng(ws.Cells(r, "C").value)
            Else
                currentVal = 0
            End If

            ws.Cells(r, "C").value = Format(currentVal + 1, "00")
            IncrementSTT_HD = ws.Cells(r, "C").value

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

' Purpose: Show/hide and highlight INPUT sheet rows based on selected collateral type
' Input:   COLLATERAL_TYPE named cell value (D28)
' Output:  Rows hidden/shown and colored accordingly

Public Sub ApplyCollateralVisibility()

    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim selectedType As String
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim applicableTypes As String
    
    Set ws = ThisWorkbook.Sheets("INPUT")
    
    ' Read selected collateral type from named cell
    selectedType = Trim(ThisWorkbook.Names("COLLATERAL_TYPE").RefersToRange.value)
    
    ' Exit if nothing selected yet
    If selectedType = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Define colors
    Dim COLOR_APPLICABLE As Long
    Dim COLOR_HEADER As Long
    Dim COLOR_NORMAL As Long
    Dim COLOR_SELECTOR As Long
    Dim COLOR_INPUT As Long
    
    COLOR_INPUT = RGB(255, 255, 153) ' Light yellow
    COLOR_APPLICABLE = RGB(198, 224, 180) ' Light green
    COLOR_HEADER = RGB(217, 217, 217)     ' Light grey
    COLOR_NORMAL = RGB(255, 255, 255)     ' White
    COLOR_SELECTOR = RGB(0, 70, 127)    ' Dark blue
    
    For i = 29 To 58
        key = Trim(ws.Cells(i, "A").value)
        applicableTypes = Trim(ws.Cells(i, "E").value)
        
        ' Selector row � always visible, always blue
        If key = "collateral_type" Then
            ws.Rows(i).Hidden = False
            ws.Rows(i).Interior.Color = COLOR_SELECTOR
            GoTo NextRow
        End If
        
        ' Section header rows (blank key) � always visible, grey
        If key = "" Then
            ws.Rows(i).Hidden = False
            ws.Rows(i).Interior.Color = COLOR_HEADER
            GoTo NextRow
        End If
        
        ' All other rows � check applicability
        If InStr(1, applicableTypes, selectedType, vbTextCompare) > 0 Then
            ' Applicable � show, highlight green, col C light yellow
            ws.Rows(i).Hidden = False
            ws.Rows(i).Interior.Color = COLOR_APPLICABLE
            ws.Cells(i, "C").Interior.Color = COLOR_INPUT
        Else
            ' Not applicable — hide and clear col C
            ws.Rows(i).Hidden = True
            ws.Rows(i).Interior.Color = COLOR_NORMAL
            ws.Cells(i, "C").ClearContents
        End If
        
NextRow:
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "L" & Chr(7895) & "i ApplyCollateralVisibility: " & Err.Description, vbCritical

End Sub
