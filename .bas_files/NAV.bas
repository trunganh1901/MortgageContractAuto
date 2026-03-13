Attribute VB_Name = "NAV"
Public Sub GoToSheet_FromC4()
    Dim targetSheet As String
    Dim ws As Worksheet

    targetSheet = Trim(ActiveSheet.Range("C4").Value)

    If targetSheet = "" Then
        MsgBox "C4 is empty.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(targetSheet)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet does not exist: " & targetSheet, vbCritical
        Exit Sub
    End If

    ' Unhide if needed
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    ws.Activate
End Sub

Public Sub GoTo_Items()
    Dim targetSheet As String
    Dim ws As Worksheet
    
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Items")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet does not exist: " & targetSheet, vbCritical
        Exit Sub
    End If
    
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    ws.Activate
End Sub

Public Sub BackTo_DashBoard()
    Dim targetSheet As String
    Dim ws As Worksheet
    
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("UI_DASHBOARD")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet does not exist: " & targetSheet, vbCritical
        Exit Sub
    End If
    
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    ws.Activate

End Sub
