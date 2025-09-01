Attribute VB_Name = "modRepositoryFactory"
Option Compare Database
Option Explicit

' =================================================================
' FACTORY: modRepositoryFactory
' RESPONSABILIDAD: Crear instancias de TODOS los repositorios.
'                  Centraliza la lógica DEV_MODE para decidir
'                  entre repositorios reales y mocks.
' PATRÓN: CERO ARGUMENTOS en métodos Create.
' =================================================================

Public Function CreateAuthRepository() As IAuthRepository
#If DEV_MODE Then
    Set CreateAuthRepository = New CMockAuthRepository
#Else
    Dim repoImpl As CAuthRepository
    Set repoImpl = New CAuthRepository
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService() ' Llama a otra factoría
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    repoImpl.Initialize config, errorHandler
    Set CreateAuthRepository = repoImpl
#End If
End Function

Public Function CreateSolicitudRepository() As ISolicitudRepository
#If DEV_MODE Then
    Set CreateSolicitudRepository = New CMockSolicitudRepository
#Else
    ' Implementación real aquí...
    Dim repoImpl As CSolicitudRepository
    Set repoImpl = New CSolicitudRepository
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    repoImpl.Initialize config, errorHandler
    Set CreateSolicitudRepository = repoImpl
#End If
End Function

Public Function CreateExpedienteRepository() As IExpedienteRepository
#If DEV_MODE Then
    Set CreateExpedienteRepository = New CMockExpedienteRepository
#Else
    Dim repoImpl As CExpedienteRepository
    Set repoImpl = New CExpedienteRepository
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    repoImpl.Initialize config, errorHandler
    Set CreateExpedienteRepository = repoImpl
#End If
End Function

Public Function CreateNotificationRepository() As INotificationRepository
#If DEV_MODE Then
    Set CreateNotificationRepository = New CMockNotificationRepository
#Else
    Dim repoImpl As CNotificationRepository
    Set repoImpl = New CNotificationRepository
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    repoImpl.Initialize config, errorHandler
    Set CreateNotificationRepository = repoImpl
#End If
End Function

Public Function CreateMapeoRepository() As IMapeoRepository
#If DEV_MODE Then
    Set CreateMapeoRepository = New CMockMapeoRepository
#Else
    Dim repoImpl As CMapeoRepository
    Set repoImpl = New CMapeoRepository
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    repoImpl.Initialize config, errorHandler
    Set CreateMapeoRepository = repoImpl
#End If
End Function

Public Function CreateWorkflowRepository() As IWorkflowRepository
#If DEV_MODE Then
    Set CreateWorkflowRepository = New CMockWorkflowRepository
#Else
    Dim repoImpl As CWorkflowRepository
    Set repoImpl = New CWorkflowRepository
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    repoImpl.Initialize config, errorHandler
    Set CreateWorkflowRepository = repoImpl
#End If
End Function

Public Function CreateOperationRepository() As IOperationRepository
#If DEV_MODE Then
    Set CreateOperationRepository = New CMockOperationRepository
#Else
    Dim repoImpl As COperationRepository
    Set repoImpl = New COperationRepository
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    repoImpl.Initialize config, errorHandler
    Set CreateOperationRepository = repoImpl
#End If
End Function

' ... Y así sucesivamente para CADA repositorio del sistema ...
' (CreateExpedienteRepository, CreateWorkflowRepository, etc.)