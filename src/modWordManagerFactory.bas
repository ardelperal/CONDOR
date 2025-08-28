Attribute VB_Name = "modWordManagerFactory"
Option Compare Database
Option Explicit


' =====================================================
' MÓDULO: modWordManagerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de gestión de Word
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-15
' =====================================================

Public Function CreateWordManager() As IWordManager
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim configService As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim fileSystem As IFileSystem
    
    ' Crear dependencias internamente
    Set configService = modConfig.CreateConfigService()
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(configService, fileSystem)
    
    ' Decidir si usar mock o clase concreta basado en DEV_MODE
    If CBool(configService.GetValue("DEV_MODE")) Then
        ' Modo desarrollo - crear instancia de Word dentro del factory
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        wordApp.DisplayAlerts = False
        
        Dim mockWordManager As New CMockWordManager
        Set CreateWordManager = mockWordManager
    Else
        ' Modo producción - crear instancia de Word y luego inicializar
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        wordApp.DisplayAlerts = False
        
        Dim wordManagerInstance As New CWordManager
        wordManagerInstance.Initialize wordApp, errorHandler
        Set CreateWordManager = wordManagerInstance
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modWordManagerFactory.CreateWordManager: " & Err.Description
    Err.Raise Err.Number, "modWordManagerFactory.CreateWordManager", Err.Description
End Function

