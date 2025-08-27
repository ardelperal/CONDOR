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


' Clase: CMockOperationRepository
' Descripción: Mock del repositorio de operaciones para pruebas unitarias.
' Arquitectura: Capa de Pruebas - Mock para aislamiento de tests

Implements IOperationRepository

' Propiedades públicas para verificación en tests
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

' Implementación del contrato IOperationRepository
Private Sub IOperationRepository_SaveLog(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    ' Registrar que el mÃ©todo fue llamado
    SaveLogCalled = True
    CallCount = CallCount + 1
    
    ' Capturar los parámetros para verificación
    LastOperationType = operationType
    LastEntityId = entityId
    LastDetails = details
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para SaveLog
Public Sub SaveLog(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    IOperationRepository_SaveLog operationType, entityId, details
End Sub


