Attribute VB_Name = "ModAppManagerFactory"
Option Compare Database
Option Explicit


Public Function CreateAppManager() As IAppManager
    On Error GoTo errorHandler
    
    Dim appManagerImpl As New CAppManager
    
    ' Crear dependencias usando sus respectivas factorías
    Dim authSvc As IAuthService
    Set authSvc = modAuthFactory.CreateAuthService()
    
    Dim configSvc As IConfig
    Set configSvc = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Inyectar dependencias
    appManagerImpl.Initialize authSvc, configSvc, errorHandler
    
    Set CreateAppManager = appManagerImpl
    
    Exit Function
    
errorHandler:
    Debug.Print "Error fatal en ModAppManagerFactory.CreateAppManager: " & Err.Description
    Set CreateAppManager = Nothing
End Function

