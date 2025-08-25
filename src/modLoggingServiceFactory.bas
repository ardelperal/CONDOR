Attribute VB_Name = "modLoggingServiceFactory"
'******************************************************************************
' Módulo: modLoggingServiceFactory
' Propósito: Factory para crear instancias de ILoggingService con inyección de dependencias
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' FUNCIONES PÚBLICAS
'******************************************************************************

' Crea una instancia completamente configurada de ILoggingService
' Retorna: Instancia de ILoggingService lista para usar
Public Function CreateLoggingService() As ILoggingService
    On Error GoTo ErrorHandler
    
    Dim service As CLoggingService
    Dim config As IConfig
    Dim fileSystem As IFileSystem
    
    ' Crear las dependencias
    Set config = modConfig.CreateConfigService()
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    ' Crear el servicio
    Set service = New CLoggingService
    
    ' Inyectar las dependencias
    service.Initialize config, fileSystem
    
    ' Devolver la instancia
    Set CreateLoggingService = service
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modLoggingServiceFactory.CreateLoggingService"
    Set CreateLoggingService = Nothing
End Function