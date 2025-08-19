VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestOptionalParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

' Test class for optional parameters
Public Function TestFunction( _
    ByVal param1 As String, _
    ByVal param2 As String, _
    Optional ByVal param3 As String = "default" _
) As Boolean
End Function