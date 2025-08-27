Attribute VB_Name = "modErrorHandlerFactory"
Option Compare Database
Option Explicit

' Módulo: modErrorHandlerFactory
' Descripción: Factory para crear instancias de IErrorHandlerService
' Arquitectura: Capa de Servicios - Factory Pattern

' Crea una instancia configurada de IErrorHandlerService
' @param config IConfig: Instancia de configuración
' @param fileSystem IFileSystem: Instancia del sistema de ficheros
' @return IErrorHandlerService: Instancia lista para usar
Public Function CreateErrorHandlerService(ByVal config As IConfig, ByVal fileSystem As IFileSystem) As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim errorHandlerInstance As New CErrorHandlerService
    
    ' Inicializar el servicio con las dependencias recibidas
    errorHandlerInstance.Initialize config, fileSystem
    
    ' Devolver la interfaz
    Set CreateErrorHandlerService = errorHandlerInstance
    
    Exit Function
    
ErrorHandler:
    ' Nota: En caso de error en el factory, usamos logging directo para evitar recursión
    Debug.Print "Error en modErrorHandlerFactory.CreateErrorHandlerService: " & Err.Number & " - " & Err.Description
    Set CreateErrorHandlerService = Nothing
End Function


