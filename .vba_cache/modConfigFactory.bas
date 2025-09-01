Attribute VB_Name = "modConfigFactory"
Option Compare Database
Option Explicit

Public Function CreateConfigService() As IConfig
    On Error GoTo ErrorHandler
    
    Dim configImpl As New CConfig
    configImpl.LoadConfiguration
    Set CreateConfigService = configImpl
    Exit Function
    
ErrorHandler:
    MsgBox "Error CRÍTICO al cargar la configuración: " & Err.Description, vbCritical, "Fallo de Arranque de CONDOR"
    Set CreateConfigService = Nothing
End Function

