WScript.Echo "Test simple"
On Error GoTo TestFail
WScript.Echo "After On Error"
WScript.Quit

TestFail:
WScript.Echo "Error: " & Err.Description