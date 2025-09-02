Attribute VB_Name = "modErrorHandlerFactory"
Option Compare Database
Option Explicit


' =====================================================
' MÓDULO: modErrorHandlerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de errores.
' PATRÓN: CERO ARGUMENTOS (Lección 37)
' =====================================================

Public Function CreateErrorHandlerService(Optional ByVal configService As IConfig = Nothing) As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim config As IConfig
    If configService Is Nothing Then
        Set config = modConfigFactory.CreateConfigService()
    Else
        Set config = configService
    End If
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem(config)
    
    Dim errorHandlerImpl As New CErrorHandlerService
    errorHandlerImpl.Initialize config, fs
    
    Set CreateErrorHandlerService = errorHandlerImpl
    Exit Function
    
ErrorHandler:
    Debug.Print "Error crítico en modErrorHandlerFactory.CreateErrorHandlerService: " & Err.Description
    Set CreateErrorHandlerService = Nothing
End Function




