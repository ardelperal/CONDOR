Option Compare Database
Option Explicit

' Variable privada para almacenar el mock
Private m_MockRepository As ISolicitudRepository

'******************************************************************************
' MÓDULO: modRepositoryFactory
' DESCRIPCIÓN: Factory para la inyección de dependencias del repositorio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************


'******************************************************************************
' FACTORY METHODS
'******************************************************************************

'******************************************************************************
' FUNCIÓN: CreateSolicitudRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de solicitudes según el modo
' RETORNA: ISolicitudRepository - Instancia del repositorio (Mock o Real)
'******************************************************************************
Public Function CreateSolicitudRepository() As ISolicitudRepository
    On Error GoTo ErrorHandler
    
    ' Si hay un mock configurado, devolverlo
    If Not m_MockRepository Is Nothing Then
        Set CreateSolicitudRepository = m_MockRepository
        Exit Function
    End If
    
    ' Obtener las dependencias
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Usar el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CSolicitudRepository
    
    ' Inyectar AMBAS dependencias
    repositoryInstance.Initialize configService, operationLogger
    
    Set CreateSolicitudRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateSolicitudRepository"
    Set CreateSolicitudRepository = Nothing
End Function

'******************************************************************************
' FUNCIÓN: CreateExpedienteRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de expedientes
' RETORNA: IExpedienteRepository - Instancia del repositorio de expedientes
'******************************************************************************
Public Function CreateExpedienteRepository() As IExpedienteRepository
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CExpedienteRepository
    
    ' Inyectar las dependencias
    repositoryInstance.Initialize configService, operationLogger
    
    Set CreateExpedienteRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateExpedienteRepository"
    Set CreateExpedienteRepository = Nothing
End Function

'******************************************************************************
' FUNCIÓN: CreateAuthRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de autenticación
' RETORNA: IAuthRepository - Instancia del repositorio de autenticación
'******************************************************************************
Public Function CreateAuthRepository() As IAuthRepository
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Crear el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CAuthRepository
    
    ' Inyectar AMBAS dependencias
    repositoryInstance.Initialize configService, operationLogger
    
    Set CreateAuthRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateAuthRepository"
    Set CreateAuthRepository = Nothing
End Function

'******************************************************************************
' FUNCIÓN: CreateMapeoRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de mapeo
' RETORNA: IMapeoRepository - Instancia del repositorio de mapeo
'******************************************************************************
Public Function CreateMapeoRepository() As IMapeoRepository
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    
    ' Crear el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CMapeoRepository
    
    ' Inyectar las dependencias
    repositoryInstance.Initialize configService
    
    Set CreateMapeoRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateMapeoRepository"
    Set CreateMapeoRepository = Nothing
End Function

'******************************************************************************
' GESTIÓN DE MOCKS PARA PRUEBAS
'******************************************************************************

'******************************************************************************
' FUNCIÓN: SetMockRepository
' DESCRIPCIÓN: Configura un mock para ser usado en lugar del repositorio real
' PARÁMETROS: mockRepo - Instancia del mock a usar
'******************************************************************************
Public Sub SetMockRepository(mockRepo As ISolicitudRepository)
    Set m_MockRepository = mockRepo
End Sub

'******************************************************************************
' FUNCIÓN: ResetMock
' DESCRIPCIÓN: Limpia el mock configurado, volviendo al comportamiento normal
'******************************************************************************
Public Sub ResetMock()
    Set m_MockRepository = Nothing
End Sub






