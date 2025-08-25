Option Compare Database
Option Explicit

' Clase: CMockOperationRepository
' DescripciÃ³n: Mock del repositorio de operaciones para pruebas unitarias.
' Arquitectura: Capa de Pruebas - Mock para aislamiento de tests

Implements IOperationRepository

' Propiedades pÃºblicas para verificaciÃ³n en tests
Public SaveLogCalled As Boolean
Public LastOperationType As String
Public LastEntityId As String
Public LastDetails As String
Public CallCount As Long

' MÃ©todo para resetear el estado del mock
Public Sub Reset()
    SaveLogCalled = False
    LastOperationType = ""
    LastEntityId = ""
    LastDetails = ""
    CallCount = 0
End Sub

' ImplementaciÃ³n del contrato IOperationRepository
Private Sub IOperationRepository_SaveLog(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    ' Registrar que el mÃ©todo fue llamado
    SaveLogCalled = True
    CallCount = CallCount + 1
    
    ' Capturar los parÃ¡metros para verificaciÃ³n
    LastOperationType = operationType
    LastEntityId = entityId
    LastDetails = details
End Sub
