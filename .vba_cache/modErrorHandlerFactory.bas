Attribute VB_Name = "modErrorHandlerFactory"
Option Compare Database
Option Explicit
' Módulo: modErrorHandlerFactory
' Descripción: Factory para crear instancias de IErrorHandlerService
' Arquitectura: Capa de Servicios - Factory Pattern

' Crea una instancia configurada de IErrorHandlerService
' @return IErrorHandlerService: Instancia lista para usar
Public Function CreateErrorHandlerService() As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim errorHandlerInstance As New CErrorHandlerService
    Dim config As IConfig
    Dim fileSystem As IFileSystem
    
    ' Obtener instancia de configuración usando el nuevo factory
    Set config = modConfig.CreateConfigService()
    
    ' Crear instancia del sistema de ficheros
    Set fileSystem = New CFileSystem
    
    ' Inicializar el servicio con sus dependencias
    errorHandlerInstance.Initialize config, fileSystem
    
    ' Devolver la interfaz
    Set CreateErrorHandlerService = errorHandlerInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modErrorHandlerFactory.CreateErrorHandlerService")
    Set CreateErrorHandlerService = Nothing
End Function