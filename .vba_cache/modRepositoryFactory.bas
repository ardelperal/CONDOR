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

Public Function CreateAuthRepository() As IAuthRepository
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repoImpl As New CAuthRepository
    repoImpl.Initialize config, errorHandler
    Set CreateAuthRepository = repoImpl
    Exit Function
errorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateAuthRepository: " & Err.Description
End Function

Public Function CreateExpedienteRepository() As IExpedienteRepository
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repoImpl As New CExpedienteRepository
    repoImpl.Initialize config, errorHandler
    Set CreateExpedienteRepository = repoImpl
    Exit Function
errorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateExpedienteRepository: " & Err.Description
End Function

Public Function CreateMapeoRepository() As IMapeoRepository
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repoImpl As New CMapeoRepository
    repoImpl.Initialize config, errorHandler
    Set CreateMapeoRepository = repoImpl
    Exit Function
errorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateMapeoRepository: " & Err.Description
End Function

Public Function CreateNotificationRepository() As INotificationRepository
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repoImpl As New CNotificationRepository
    repoImpl.Initialize config, errorHandler
    Set CreateNotificationRepository = repoImpl
    Exit Function
errorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateNotificationRepository: " & Err.Description
End Function

Public Function CreateOperationRepository() As IOperationRepository
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repoImpl As New COperationRepository
    repoImpl.Initialize config, errorHandler
    Set CreateOperationRepository = repoImpl
    Exit Function
errorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateOperationRepository: " & Err.Description
End Function

Public Function CreateSolicitudRepository() As ISolicitudRepository
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repoImpl As New CSolicitudRepository
    repoImpl.Initialize config, errorHandler
    Set CreateSolicitudRepository = repoImpl
    Exit Function
errorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateSolicitudRepository: " & Err.Description
End Function

Public Function CreateWorkflowRepository() As IWorkflowRepository
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repoImpl As New CWorkflowRepository
    repoImpl.Initialize config, errorHandler
    Set CreateWorkflowRepository = repoImpl
    Exit Function
errorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateWorkflowRepository: " & Err.Description
End Function

