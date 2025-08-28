Attribute VB_Name = "modRepositoryFactory"
Option Compare Database
Option Explicit

' ... (Encabezado del módulo) ...

' Patrón Estándar: Firma (config, errorHandler), con chequeo de DEV_MODE

Public Function CreateSolicitudRepository(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService) As ISolicitudRepository
    If CBool(config.GetValue("DEV_MODE")) Then
        Set CreateSolicitudRepository = New CMockSolicitudRepository
    Else
        Dim repo As New CSolicitudRepository
        repo.Initialize config
        Set CreateSolicitudRepository = repo
    End If
End Function

Public Function CreateExpedienteRepository(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService) As IExpedienteRepository
    If CBool(config.GetValue("DEV_MODE")) Then
        Set CreateExpedienteRepository = New CMockExpedienteRepository
    Else
        Dim repo As New CExpedienteRepository
        repo.Initialize config, errorHandler
        Set CreateExpedienteRepository = repo
    End If
End Function

Public Function CreateAuthRepository(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService) As IAuthRepository
    If CBool(config.GetValue("DEV_MODE")) Then
        Set CreateAuthRepository = New CMockAuthRepository
    Else
        Dim repo As New CAuthRepository
        repo.Initialize config, errorHandler
        Set CreateAuthRepository = repo
    End If
End Function

Public Function CreateNotificationRepository(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService) As INotificationRepository
    If CBool(config.GetValue("DEV_MODE")) Then
        Set CreateNotificationRepository = New CMockNotificationRepository
    Else
        Dim repo As New CNotificationRepository
        repo.Initialize config, errorHandler
        Set CreateNotificationRepository = repo
    End If
End Function

Public Function CreateMapeoRepository(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService) As IMapeoRepository
    If CBool(config.GetValue("DEV_MODE")) Then
        Set CreateMapeoRepository = New CMockMapeoRepository
    Else
        Dim repo As New CMapeoRepository
        repo.Initialize config, errorHandler
        Set CreateMapeoRepository = repo
    End If
End Function

Public Function CreateWorkflowRepository(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService) As IWorkflowRepository
    If CBool(config.GetValue("DEV_MODE")) Then
        Set CreateWorkflowRepository = New CMockWorkflowRepository
    Else
        Dim repo As New CWorkflowRepository
        repo.Initialize config, errorHandler
        Set CreateWorkflowRepository = repo
    End If
End Function

Public Function CreateOperationRepository(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService) As IOperationRepository
    If CBool(config.GetValue("DEV_MODE")) Then
        Set CreateOperationRepository = New CMockOperationRepository
    Else
        Dim repo As New COperationRepository
        repo.Initialize config, errorHandler
        Set CreateOperationRepository = repo
    End If
End Function