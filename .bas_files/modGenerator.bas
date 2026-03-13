Attribute VB_Name = "modGenerator"
Sub GeneralDocGenerate()

    Dim wbPath As String
    Dim pyCmd As String

    On Error GoTo ERR_HANDLER

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Save before Python
    ThisWorkbook.Save

    wbPath = ThisWorkbook.Path

    pyCmd = _
        "import sys;" & _
        " sys.path.insert(0, r'" & wbPath & "');" & _
        " import app.main;" & _
        " app.main.main()"

    Application.Run "RunPython", pyCmd

    IncrementSTT_HD
    ThisWorkbook.Save
    OpenCustomerOutputFolder

CLEAN_EXIT:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ERR_HANDLER:
    MsgBox "Error while generating contracts." & vbCrLf & _
           Err.Description, vbCritical, "Error"
    Resume CLEAN_EXIT

End Sub

