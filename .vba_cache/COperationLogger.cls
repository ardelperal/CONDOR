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
    On Error GoTo ErrorHandler
    Set m_configService = config
    Set m_OperationRepository = repository
    Exit Sub
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "COperationLogger.Initialize")
End Sub



' Implementación del contrato IOperationLogger
Private Sub IOperationLogger_LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    On Error GoTo ErrorHandler
    
    ' Delegar la operación de guardado al repositorio
    m_OperationRepository.SaveLog operationType, entityId, details
    
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogOperation", _
                                  "TipoOperacion: " & operationType & ", Entidad: " & entityId)
End Sub
