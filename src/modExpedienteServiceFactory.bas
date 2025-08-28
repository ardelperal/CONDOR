Attribute VB_Name = "modExpedienteServiceFactory"
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
    
    ' 1. Obtener dependencias de sus propias factorías
    Dim repo As IExpedienteRepository
    Dim logger As IOperationLogger
    Dim errorHandler As IErrorHandlerService
    Dim config As IConfig
    
    ' Crear errorHandler y config primero para poder usarlos en caso de errores
    Set config = modConfig.CreateConfigService()
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Crear las demás dependencias
    Set repo = modRepositoryFactory.CreateExpedienteRepository(config, errorHandler)
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' 2. Crear e inicializar el servicio
    Dim expedienteServiceInstance As New CExpedienteService
    expedienteServiceInstance.Initialize config, logger, repo, errorHandler
    
    ' 3. Devolver la instancia
    Set CreateExpedienteService = expedienteServiceInstance
    
    Exit Function
    
ErrorHandler:
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modExpedienteServiceFactory.CreateExpedienteService"
    End If
    Set CreateExpedienteService = Nothing
End Function

