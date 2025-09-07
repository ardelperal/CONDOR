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
Private m_ErrorHandler As IErrorHandlerService

' Método de inicialización para inyectar dependencias
Public Sub Initialize(config As IConfig, repository As IOperationRepository, errorHandler As IErrorHandlerService)
    On Error GoTo errorHandler
    
    Set m_configService = config
    Set m_OperationRepository = repository
    Set m_ErrorHandler = errorHandler
    Exit Sub
errorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.Initialize"
End Sub



' Implementación del contrato IOperationLogger
Private Sub IOperationLogger_LogOperation(ByVal operationType As String, ByVal entityId As Long, ByVal details As String)
    On Error GoTo errorHandler
    
    ' Delegar la operación de guardado al repositorio
    m_OperationRepository.SaveLog operationType, entityId, details
    
    Exit Sub
    
errorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogOperation", _
                                  "TipoOperacion: " & operationType & ", Entidad: " & entityId
End Sub

Private Sub IOperationLogger_LogSolicitudCreation(ByVal SolicitudID As Long, ByVal userId As String)
    On Error GoTo errorHandler
    
    ' Registrar la creación de solicitud
    m_OperationRepository.SaveLog "SOLICITUD_CREADA", SolicitudID, "Usuario: " & userId
    
    Exit Sub
    
errorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogSolicitudCreation", _
                                  "SolicitudId: " & SolicitudID & ", Usuario: " & userId
End Sub

Private Sub IOperationLogger_LogStateChange(ByVal SolicitudID As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
    On Error GoTo errorHandler
    
    ' Registrar el cambio de estado
    Dim details As String
    details = "Estado anterior: " & oldState & ", Estado nuevo: " & newState & ", Usuario: " & userId
    m_OperationRepository.SaveLog "CAMBIO_ESTADO", SolicitudID, details
    
    Exit Sub
    
errorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogStateChange", _
                                  "SolicitudId: " & SolicitudID & ", Estados: " & oldState & " -> " & newState
End Sub

Private Sub IOperationLogger_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
    On Error GoTo errorHandler
    
    ' Registrar el error del sistema
    Dim details As String
    details = "Número: " & errNumber & ", Descripción: " & errDescription & ", Módulo: " & moduleName & ", Crítico: " & isCritical
    m_OperationRepository.SaveLog "ERROR_SISTEMA", errNumber, details
    
    Exit Sub
    
errorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogError", _
                                  "Número: " & errNumber & ", Módulo: " & moduleName
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Sub LogOperation(ByVal operationType As String, ByVal entityId As Long, ByVal details As String)
    IOperationLogger_LogOperation operationType, entityId, details
End Sub

Public Sub LogSolicitudCreation(ByVal SolicitudID As Long, ByVal userId As String)
    IOperationLogger_LogSolicitudCreation SolicitudID, userId
End Sub

Public Sub LogStateChange(ByVal SolicitudID As Long, ByVal oldState As String, ByVal newState As String, ByVal userId As String)
    IOperationLogger_LogStateChange SolicitudID, oldState, newState, userId
End Sub

Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
    IOperationLogger_LogError errNumber, errDescription, moduleName, isCritical
End Sub






