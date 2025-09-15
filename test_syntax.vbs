Private Sub TestSub()
    WScript.Echo "Test"
    On Error GoTo TestFail
    WScript.Echo "After On Error"
    Exit Sub
TestFail:
    WScript.Echo "Error: " & Err.Description
End Sub

TestSub
