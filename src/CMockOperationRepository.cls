VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockOperationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IOperationRepository

' Variables públicas para seguimiento del mock
Public SaveLogCalled As Boolean
Public CallCount As Long
Public LastOperationType As String
Public LastEntityId As Long
Public LastDetails As String

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

Private Sub IOperationRepository_SaveLog(ByVal operationType As String, ByVal entityId As Long, ByVal details As String)
    ' Registrar la llamada para verificación en tests
    SaveLogCalled = True
    CallCount = CallCount + 1
    LastOperationType = operationType
    LastEntityId = entityId
    LastDetails = details
End Sub

' Resetea el estado del mock para asegurar el aislamiento entre tests
Public Sub Reset()
    SaveLogCalled = False
    CallCount = 0
    LastOperationType = ""
    LastEntityId = 0
    LastDetails = ""
End Sub

