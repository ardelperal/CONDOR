VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockOperationLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


#If DEV_MODE Then

' Clase: CMockOperationLogger
' Descripción: Mock del servicio de logging de operaciones para pruebas unitarias.
'              Almacena las operaciones loggeadas en una colección para su verificación.

Implements IOperationLogger

Public LoggedOperations As Collection
Public LogOperationCalled As Boolean

Private Sub Class_Initialize()
    Set LoggedOperations = New Collection
    LogOperationCalled = False
End Sub

Private Sub IOperationLogger_LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Registrar que el método fue llamado
    LogOperationCalled = True
    
    Dim logEntry As New Collection
    logEntry.Add operationType, "OperationType"
    logEntry.Add entityId, "EntityId"
    logEntry.Add details, "Details"
    logEntry.Add "LogOperation", "Method"
    
    LoggedOperations.Add logEntry
    
    Exit Sub
    
ErrorHandler:
    ' En un mock, los errores internos no suelen ser críticos, pero se podrían registrar para depuración del mock
    Debug.Print "Error en CMockOperationLogger.IOperationLogger_LogOperation: " & Err.Description
End Sub

Private Sub IOperationLogger_LogSolicitudCreation(ByVal SolicitudID As Long, ByVal userId As String)
    On Error GoTo ErrorHandler
    
    ' Registrar que el método fue llamado
    LogOperationCalled = True
    
    Dim logEntry As New Collection
    logEntry.Add "SOLICITUD_CREADA", "OperationType"
    logEntry.Add CStr(SolicitudID), "EntityId"
    logEntry.Add "Usuario: " & userId, "Details"
    logEntry.Add "LogSolicitudCreation", "Method"
    
    LoggedOperations.Add logEntry
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CMockOperationLogger.IOperationLogger_LogSolicitudCreation: " & Err.Description
End Sub

Private Sub IOperationLogger_LogStateChange(ByVal SolicitudID As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
    On Error GoTo ErrorHandler
    
    ' Registrar que el método fue llamado
    LogOperationCalled = True
    
    Dim logEntry As New Collection
    logEntry.Add "CAMBIO_ESTADO", "OperationType"
    logEntry.Add CStr(SolicitudID), "EntityId"
    logEntry.Add "Estado anterior: " & oldState & ", Estado nuevo: " & newState & ", Usuario: " & userId, "Details"
    logEntry.Add "LogStateChange", "Method"
    
    LoggedOperations.Add logEntry
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CMockOperationLogger.IOperationLogger_LogStateChange: " & Err.Description
End Sub

Private Sub IOperationLogger_LogError(ByVal errorCode As String, ByVal errorMessage As String, ByVal context As String)
    On Error GoTo ErrorHandler
    
    ' Registrar que el método fue llamado
    LogOperationCalled = True
    
    Dim logEntry As New Collection
    logEntry.Add "ERROR_SISTEMA", "OperationType"
    logEntry.Add errorCode, "EntityId"
    logEntry.Add "Código: " & errorCode & ", Mensaje: " & errorMessage & ", Contexto: " & context, "Details"
    logEntry.Add "LogError", "Method"
    
    LoggedOperations.Add logEntry
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CMockOperationLogger.IOperationLogger_LogError: " & Err.Description
End Sub

Public Sub Reset()
    Set LoggedOperations = New Collection
    LogOperationCalled = False
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para LogOperation
Public Sub LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    IOperationLogger_LogOperation operationType, entityId, details
End Sub

' Método público de conveniencia para LogSolicitudCreation
Public Sub LogSolicitudCreation(ByVal SolicitudID As Long, ByVal userId As String)
    IOperationLogger_LogSolicitudCreation SolicitudID, userId
End Sub

' Método público de conveniencia para LogStateChange
Public Sub LogStateChange(ByVal SolicitudID As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
    IOperationLogger_LogStateChange SolicitudID, oldState, newState, userId
End Sub

' Método público de conveniencia para LogError
Public Sub LogError(ByVal errorCode As String, ByVal errorMessage As String, ByVal context As String)
    IOperationLogger_LogError errorCode, errorMessage, context
End Sub

#End If


