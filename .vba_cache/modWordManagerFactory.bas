Attribute VB_Name = "modWordManagerFactory"
Option Compare Database
Option Explicit


' =====================================================
' MÓDULO: modWordManagerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de gestión de Word
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-15
' =====================================================

Public Function CreateWordManager(ByVal errorHandler As IErrorHandlerService) As IWordManager
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim configService As IConfig
    
    ' Obtener servicio de configuración
    Set configService = modConfig.CreateConfigService(, errorHandler)
    
    ' Decidir si usar mock o clase concreta basado en DEV_MODE
    If CBool(configService.GetValue("DEV_MODE")) Then
        ' Modo desarrollo - crear instancia de Word dentro del factory
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        wordApp.DisplayAlerts = False
        
        Dim mockWordManager As New CMockWordManager
        mockWordManager.Initialize wordApp, errorHandler
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
    errorHandler.LogError Err.Number, Err.Description, "modWordManagerFactory.CreateWordManager"
    Set CreateWordManager = Nothing
End Function

