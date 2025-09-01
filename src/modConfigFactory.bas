Attribute VB_Name = "modConfigFactory"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: modConfigFactory
' DESCRIPCIÓN: Factory para la creación del servicio de configuración
' PATRÓN: CERO ARGUMENTOS (Lección 37)
' =====================================================

' --- FUNCIÓN FACTORY PRINCIPAL ---
Public Function CreateConfigService() As IConfig
    On Error GoTo ErrorHandler
    
    Dim configImpl As New CConfig
    
    ' La factoría resuelve sus propias dependencias internamente
    Dim errHandler As IErrorHandlerService
    Set errHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    configImpl.Initialize errHandler
    
    Set CreateConfigService = configImpl
    Exit Function
    
ErrorHandler:
    Debug.Print "Error crítico en modConfigFactory.CreateConfigService: " & Err.Description
    Set CreateConfigService = Nothing
End Function