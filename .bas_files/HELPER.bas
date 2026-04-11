Attribute VB_Name = "HELPER"
Option Explicit
Sub OpenCustomerOutputFolder()

    Dim folderPath As String
    Dim userchoice As VbMsgBoxResult

    folderPath = BuildPath(ThisWorkbook.Path, "Output")

    If folderPath = "" Then
        MsgBox "No output folder path found.", vbExclamation
        Exit Sub
    End If

    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder not found:" & vbCrLf & folderPath, vbExclamation
        Exit Sub
    End If
    
    ' Success + ask to open folder
    userchoice = MsgBox( _
        "Contracts generated successfully!" & vbCrLf & vbCrLf & _
        "Output folder:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
        "Open this folder now?", _
        vbQuestion + vbYesNo, _
        "Done" _
    )

    If userchoice = vbYes Then
        If folderPath <> "" And Dir(folderPath, vbDirectory) <> "" Then
            Shell "explorer.exe """ & folderPath & """", vbNormalFocus
        Else
            MsgBox "Output folder not found.", vbExclamation
        End If
    End If
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

