Attribute VB_Name = "modWordManagerFactory"
Option Compare Database
Option Explicit



' =====================================================
' MÓDULO: modWordManagerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de gestión de Word
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-15
' =====================================================

Public Function CreateWordManager(Optional ByVal config As IConfig = Nothing) As IWordManager
    On Error GoTo errorHandler
    
    ' Determinar configuración final
    Dim finalConfig As IConfig
    If config Is Nothing Then
        Set finalConfig = modConfigFactory.CreateConfigService()
    Else
        Set finalConfig = config
    End If
    
    Dim wordApp As Object
    Dim errorHandler As IErrorHandlerService
    Dim fileSystem As IFileSystem
    
    ' Crear dependencias propagando la configuración
    Set fileSystem = modFileSystemFactory.CreateFileSystem(finalConfig)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(finalConfig)
    
    ' Crear instancia de Word y luego inicializar
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    
    Dim wordManagerInstance As New CWordManager
    wordManagerInstance.Initialize wordApp, errorHandler
    Set CreateWordManager = wordManagerInstance
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modWordManagerFactory.CreateWordManager: " & Err.Description
    Err.Raise Err.Number, "modWordManagerFactory.CreateWordManager", Err.Description
End Function



