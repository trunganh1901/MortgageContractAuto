Attribute VB_Name = "ValidateTemplateFn"
Function ValidateTemplateSheet(tplSheet As Worksheet) As Boolean
    Dim r As Long
    Dim missing As String
    missing = ""

    r = 5
    Do While Trim(tplSheet.Cells(r, 1).Value) <> ""
        If tplSheet.Cells(r, 1).Font.Bold Then
            If Trim(tplSheet.Cells(r, 4).Value) = "" Then
                missing = missing & "- " & tplSheet.Cells(r, 1).Value & vbCrLf
            End If
        End If
        r = r + 1
    Loop

    If missing <> "" Then
        MsgBox "Missing required fields:" & vbCrLf & missing, vbCritical
        ValidateTemplateSheet = False
    Else
        ValidateTemplateSheet = True
    End If
End Function

