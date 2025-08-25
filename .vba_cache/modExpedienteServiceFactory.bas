Option Compare Database
Option Explicit

' =====================================================
' MODULO: modExpedienteServiceFactory
' DESCRIPCION: Factory especializada para la creación del servicio de expedientes
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

' Función factory para crear y configurar el servicio de expedientes
Public Function CreateExpedienteService() As IExpedienteService
    On Error GoTo ErrorHandler
    
    ' Obtener todas las dependencias requeridas
    Dim m_Config As IConfig
    Dim m_OperationLogger As IOperationLogger
    Dim m_ExpedienteRepository As IExpedienteRepository
    
    Set m_Config = modConfig.CreateConfigService()
    Set m_OperationLogger = modOperationLoggerFactory.CreateOperationLogger()
    Set m_ExpedienteRepository = modRepositoryFactory.CreateExpedienteRepository()
    
    ' Crear una instancia de la clase concreta
    Dim expedienteServiceInstance As New CExpedienteService
    
    ' Inicializar la instancia concreta con todas las dependencias
    expedienteServiceInstance.Initialize m_Config, m_OperationLogger, m_ExpedienteRepository
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateExpedienteService = expedienteServiceInstance
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modExpedienteServiceFactory.CreateExpedienteService"
    Set CreateExpedienteService = Nothing
End Function