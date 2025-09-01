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
    On Error GoTo errorHandler
    
    Dim wordApp As Object
    Dim configService As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim fileSystem As IFileSystem
    
    ' Crear dependencias internamente
    Set configService = modConfigFactory.CreateConfigService()
    Set fileSystem = modFileSystemFactory.CreateFileSystem
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService
    
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



