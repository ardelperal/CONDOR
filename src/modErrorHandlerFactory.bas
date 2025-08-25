Attribute VB_Name = "modErrorHandlerFactory"
Option Compare Database
Option Explicit
' MÃ³dulo: modErrorHandlerFactory
' DescripciÃ³n: Factory para crear instancias de IErrorHandlerService
' Arquitectura: Capa de Servicios - Factory Pattern

' Crea una instancia configurada de IErrorHandlerService
' @return IErrorHandlerService: Instancia lista para usar
Public Function CreateErrorHandlerService() As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim errorHandlerInstance As New CErrorHandlerService
    Dim config As IConfig
    Dim fileSystem As IFileSystem
    
    ' Obtener instancia de configuraciÃ³n usando el nuevo factory
    Set config = modConfig.CreateConfigService()
    
    ' Crear instancia del sistema de ficheros
    Set fileSystem = New CFileSystem
    
    ' Inicializar el servicio con sus dependencias
    errorHandlerInstance.Initialize config, fileSystem
    
    ' Devolver la interfaz
    Set CreateErrorHandlerService = errorHandlerInstance
    
    Exit Function
    
ErrorHandler:
    ' Nota: En caso de error en el factory, usamos logging directo para evitar recursiÃ³n
    Debug.Print "Error en modErrorHandlerFactory.CreateErrorHandlerService: " & Err.Number & " - " & Err.Description
    Set CreateErrorHandlerService = Nothing
End Function
