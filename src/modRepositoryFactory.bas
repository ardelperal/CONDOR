Attribute VB_Name = "modRepositoryFactory"
Option Compare Database
Option Explicit


' =====================================================
' FACTORY: modRepositoryFactory
' DESCRIPCIÓN: Crea instancias de TODOS los repositorios del sistema.
' PATRÓN ARQUITECTÓNICO: Factory Testeable con Parámetros Opcionales.
' ESTÁNDAR: Oro - Versión Completa
' =====================================================

' Flag para alternar entre implementaciones reales y mocks
Public Const DEV_MODE As Boolean = True

' --- Métodos de Creación de Repositorios ---

Public Function CreateAuthRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IAuthRepository
    On Error GoTo errorHandler
    Dim repoImpl As New CAuthRepository
    repoImpl.Initialize GetEffectiveConfig(config), GetEffectiveErrorHandler(errorHandler)
    Set CreateAuthRepository = repoImpl
    Exit Function
errorHandler:
    HandleFactoryError "CreateAuthRepository"
End Function

Public Function CreateExpedienteRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IExpedienteRepository
    On Error GoTo errorHandler
    Dim repoImpl As New CExpedienteRepository
    repoImpl.Initialize GetEffectiveConfig(config), GetEffectiveErrorHandler(errorHandler)
    Set CreateExpedienteRepository = repoImpl
    Exit Function
errorHandler:
    HandleFactoryError "CreateExpedienteRepository"
End Function

Public Function CreateMapeoRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IMapeoRepository
    On Error GoTo errorHandler
    Dim repoImpl As New CMapeoRepository
    repoImpl.Initialize GetEffectiveConfig(config), GetEffectiveErrorHandler(errorHandler)
    Set CreateMapeoRepository = repoImpl
    Exit Function
errorHandler:
    HandleFactoryError "CreateMapeoRepository"
End Function

Public Function CreateNotificationRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As INotificationRepository
    On Error GoTo errorHandler
    Dim repoImpl As New CNotificationRepository
    repoImpl.Initialize GetEffectiveConfig(config), GetEffectiveErrorHandler(errorHandler)
    Set CreateNotificationRepository = repoImpl
    Exit Function
errorHandler:
    HandleFactoryError "CreateNotificationRepository"
End Function

Public Function CreateOperationRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IOperationRepository
    On Error GoTo errorHandler
    Dim repoImpl As New COperationRepository
    repoImpl.Initialize GetEffectiveConfig(config), GetEffectiveErrorHandler(errorHandler)
    Set CreateOperationRepository = repoImpl
    Exit Function
errorHandler:
    HandleFactoryError "CreateOperationRepository"
End Function

Public Function CreateSolicitudRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As ISolicitudRepository
    On Error GoTo errorHandler
    Dim repoImpl As New CSolicitudRepository
    repoImpl.Initialize GetEffectiveConfig(config), GetEffectiveErrorHandler(errorHandler)
    Set CreateSolicitudRepository = repoImpl
    Exit Function
errorHandler:
    HandleFactoryError "CreateSolicitudRepository"
End Function

Public Function CreateWorkflowRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IWorkflowRepository
    On Error GoTo errorHandler
    Dim repoImpl As New CWorkflowRepository
    repoImpl.Initialize GetEffectiveConfig(config), GetEffectiveErrorHandler(errorHandler)
    Set CreateWorkflowRepository = repoImpl
    Exit Function
errorHandler:
    HandleFactoryError "CreateWorkflowRepository"
End Function

' --- Funciones Auxiliares Privadas ---

Private Function GetEffectiveConfig(Optional ByVal config As IConfig = Nothing) As IConfig
    If config Is Nothing Then
        Set GetEffectiveConfig = modTestContext.GetTestConfig()
    Else
        Set GetEffectiveConfig = config
    End If
End Function

Private Function GetEffectiveErrorHandler(Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IErrorHandlerService
    If errorHandler Is Nothing Then
        Set GetEffectiveErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Else
        Set GetEffectiveErrorHandler = errorHandler
    End If
End Function

Private Sub HandleFactoryError(ByVal functionName As String)
    Debug.Print "Error crítico en modRepositoryFactory." & functionName & ": " & Err.Description
End Sub

