Option Compare Database
Option Explicit

' Clase: CMockOperationLogger
' Descripción: Mock del servicio de logging de operaciones para pruebas unitarias.
'              Almacena las operaciones loggeadas en una colección para su verificación.

Implements IOperationLogger

Public LoggedOperations As Collection

Private Sub Class_Initialize()
    Set LoggedOperations = New Collection
End Sub

Private Sub IOperationLogger_LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    On Error GoTo ErrorHandler
    
    Dim logEntry As New Collection
    logEntry.Add operationType, "OperationType"
    logEntry.Add entityId, "EntityId"
    logEntry.Add details, "Details"
    
    LoggedOperations.Add logEntry
    
    Exit Sub
    
ErrorHandler:
    ' En un mock, los errores internos no suelen ser críticos, pero se podrían registrar para depuración del mock
    Debug.Print "Error en CMockOperationLogger.IOperationLogger_LogOperation: " & Err.Description
End Sub

Public Sub ClearLog()
    Set LoggedOperations = New Collection
End Sub
