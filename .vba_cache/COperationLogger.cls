Option Compare Database
Option Explicit

' Clase: COperationLogger
' DescripciÃ³n: Servicio para el registro de operaciones del sistema.
' Arquitectura: Capa de LÃ³gica de Negocio - ImplementaciÃ³n de IOperationLogger

Implements IOperationLogger

' Dependencias inyectadas
Private m_configService As IConfig
Private m_OperationRepository As IOperationRepository

' MÃ©todo de inicializaciÃ³n para inyectar dependencias
Public Sub Initialize(config As IConfig, repository As IOperationRepository)
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    
    Set m_configService = config
    Set m_OperationRepository = repository
    Exit Sub
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "COperationLogger.Initialize"
End Sub



' ImplementaciÃ³n del contrato IOperationLogger
Private Sub IOperationLogger_LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    
    ' Delegar la operaciÃ³n de guardado al repositorio
    m_OperationRepository.SaveLog operationType, entityId, details
    
    Exit Sub
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogOperation", _
                                  "TipoOperacion: " & operationType & ", Entidad: " & entityId
End Sub
