Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Name As String
Public Passed As Boolean
Public ErrorMessage As String

Public Sub Initialize(ByVal testName As String)
    Name = testName
    Passed = False
    ErrorMessage = ""
End Sub

Public Sub Pass()
    Passed = True
    ErrorMessage = ""
End Sub

Public Sub Fail(Optional ByVal message As String = "Test failed without a specific message.")
    Passed = False
    ErrorMessage = message
<<<<<<< HEAD
End Sub
=======
End Sub






>>>>>>> main
