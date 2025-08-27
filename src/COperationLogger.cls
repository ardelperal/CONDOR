VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COperationLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Clase: COperationLogger
' Descripción: Servicio para el registro de operaciones del sistema.
' Arquitectura: Capa de Lógica de Negocio - Implementación de IOperationLogger

Implements IOperationLogger

' Dependencias inyectadas
Private m_configService As IConfig
Private m_OperationRepository As IOperationRepository

' Método de inicialización para inyectar dependencias
Public Sub Initialize(config As IConfig, repository As IOperationRepository)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    
    Set m_configService = config
    Set m_OperationRepository = repository
    Exit Sub
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.Initialize"
End Sub



' Implementación del contrato IOperationLogger
Private Sub IOperationLogger_LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    
    ' Delegar la operación de guardado al repositorio
    m_OperationRepository.SaveLog operationType, entityId, details
    
    Exit Sub
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogOperation", _
                                  "TipoOperacion: " & operationType & ", Entidad: " & entityId
End Sub

Private Sub IOperationLogger_LogSolicitudCreation(ByVal SolicitudID As Long, ByVal userId As String)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Registrar la creación de solicitud
    m_OperationRepository.SaveLog "SOLICITUD_CREADA", CStr(SolicitudID), "Usuario: " & userId
    
    Exit Sub
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogSolicitudCreation", _
                                  "SolicitudId: " & SolicitudID & ", Usuario: " & userId
End Sub

Private Sub IOperationLogger_LogStateChange(ByVal SolicitudID As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Registrar el cambio de estado
    Dim details As String
    details = "Estado anterior: " & oldState & ", Estado nuevo: " & newState & ", Usuario: " & userId
    m_OperationRepository.SaveLog "CAMBIO_ESTADO", CStr(SolicitudID), details
    
    Exit Sub
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogStateChange", _
                                  "SolicitudId: " & SolicitudID & ", Estados: " & oldState & " -> " & newState
End Sub

Private Sub IOperationLogger_LogError(ByVal errorCode As String, ByVal errorMessage As String, ByVal context As String)
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Registrar el error del sistema
    Dim details As String
    details = "Código: " & errorCode & ", Mensaje: " & errorMessage & ", Contexto: " & context
    m_OperationRepository.SaveLog "ERROR_SISTEMA", errorCode, details
    
    Exit Sub
    
ErrorHandler:
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogError", _
                                  "Código: " & errorCode & ", Contexto: " & context
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Sub LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    IOperationLogger_LogOperation operationType, entityId, details
End Sub

Public Sub LogSolicitudCreation(ByVal SolicitudID As Long, ByVal userId As String)
    IOperationLogger_LogSolicitudCreation SolicitudID, userId
End Sub

Public Sub LogStateChange(ByVal SolicitudID As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
    IOperationLogger_LogStateChange SolicitudID, oldState, newState, userId
End Sub

Public Sub LogError(ByVal errorCode As String, ByVal errorMessage As String, ByVal context As String)
    IOperationLogger_LogError errorCode, errorMessage, context
End Sub



