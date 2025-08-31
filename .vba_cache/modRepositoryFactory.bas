Attribute VB_Name = "modRepositoryFactory"
Option Compare Database
Option Explicit

' ... (Encabezado del módulo) ...

' Patrón Estándar: Firma (config, errorHandler)

Public Function CreateSolicitudRepository() As ISolicitudRepository
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repo As New CSolicitudRepository
    repo.Initialize config, errorHandler
    Set CreateSolicitudRepository = repo
End Function

Public Function CreateExpedienteRepository() As IExpedienteRepository
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repo As New CExpedienteRepository
    repo.Initialize config, errorHandler
    Set CreateExpedienteRepository = repo
End Function

Public Function CreateAuthRepository() As IAuthRepository
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repo As New CAuthRepository
    repo.Initialize config, errorHandler
    Set CreateAuthRepository = repo
End Function

Public Function CreateNotificationRepository() As INotificationRepository
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repo As New CNotificationRepository
    repo.Initialize config, errorHandler
    Set CreateNotificationRepository = repo
End Function

Public Function CreateMapeoRepository() As IMapeoRepository
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repo As New CMapeoRepository
    repo.Initialize config, errorHandler
    Set CreateMapeoRepository = repo
End Function

Public Function CreateWorkflowRepository() As IWorkflowRepository
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repo As New CWorkflowRepository
    repo.Initialize config, errorHandler
    Set CreateWorkflowRepository = repo
End Function

Public Function CreateOperationRepository() As IOperationRepository
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repo As New COperationRepository
    repo.Initialize config, errorHandler
    Set CreateOperationRepository = repo
End Function