Attribute VB_Name = "modRepositoryFactory"
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
Public Function CreateSolicitudRepository(ByVal errorHandler As IErrorHandlerService, ByVal configService As IConfig) As ISolicitudRepository
    On Error GoTo ErrorHandler
    
    ' Si hay un mock configurado, devolverlo
    If Not m_MockRepository Is Nothing Then
        Set CreateSolicitudRepository = m_MockRepository
        Exit Function
    End If
    
    ' Decidir si usar mock o clase concreta basado en DEV_MODE
    If CBool(configService.GetValue("DEV_MODE")) Then
        ' Modo desarrollo - usar mock
        Dim mockRepository As New CMockSolicitudRepository
        mockRepository.Initialize errorHandler
        Set CreateSolicitudRepository = mockRepository
    Else
        ' Modo producción - usar clase concreta
        Dim repositoryInstance As New CSolicitudRepository
        repositoryInstance.Initialize configService
        Set CreateSolicitudRepository = repositoryInstance
    End If
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateSolicitudRepository"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'******************************************************************************
' FUNCIÓN: CreateExpedienteRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de expedientes
' RETORNA: IExpedienteRepository - Instancia del repositorio de expedientes
'******************************************************************************
Public Function CreateExpedienteRepository(ByVal errorHandler As IErrorHandlerService) As IExpedienteRepository
    On Error GoTo ErrorHandler
    
    ' Obtener las dependencias
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService(errorHandler)
    
    ' Crear el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CExpedienteRepository
    
    ' Inyectar dependencias
    repositoryInstance.Initialize configService, errorHandler
    
    Set CreateExpedienteRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateExpedienteRepository"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'******************************************************************************
' FUNCIÓN: CreateAuthRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de autenticación
' RETORNA: IAuthRepository - Instancia del repositorio de autenticación
'******************************************************************************
Public Function CreateAuthRepository(ByVal errorHandler As IErrorHandlerService, ByVal configService As IConfig) As IAuthRepository
    On Error GoTo ErrorHandler
    
    ' Crear el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CAuthRepository
    
    ' Inyectar dependencias
    repositoryInstance.Initialize configService, errorHandler
    
    Set CreateAuthRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateAuthRepository"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'******************************************************************************
' FUNCIÓN: CreateNotificationRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de notificaciones
' RETORNA: INotificationRepository - Instancia del repositorio de notificaciones
'******************************************************************************
Public Function CreateNotificationRepository(ByVal errorHandler As IErrorHandlerService, ByVal configService As IConfig) As INotificationRepository
    On Error GoTo ErrorHandler
    
    ' Crear el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CNotificationRepository
    
    ' Inyectar dependencias
    repositoryInstance.Initialize configService, errorHandler
    
    Set CreateNotificationRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateNotificationRepository"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'******************************************************************************
' FUNCIÓN: CreateMapeoRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de mapeo
' RETORNA: IMapeoRepository - Instancia del repositorio de mapeo
'******************************************************************************
Public Function CreateMapeoRepository(ByVal configService As IConfig, ByVal errorHandler As IErrorHandlerService) As IMapeoRepository
    On Error GoTo ErrorHandler
    
    ' Decidir si usar mock o clase concreta basado en DEV_MODE
    If CBool(configService.GetValue("DEV_MODE")) Then
        ' Modo desarrollo - usar mock
        Dim mockRepository As New CMockMapeoRepository
        mockRepository.Initialize errorHandler
        Set CreateMapeoRepository = mockRepository
    Else
        ' Modo producción - usar clase concreta
        Dim repositoryInstance As New CMapeoRepository
        repositoryInstance.Initialize configService, errorHandler
        Set CreateMapeoRepository = repositoryInstance
    End If
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modRepositoryFactory.CreateMapeoRepository"
    Err.Raise Err.Number, Err.Source, Err.Description
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
    On Error GoTo ErrorHandler
    Set m_MockRepository = mockRepo
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modRepositoryFactory.SetMockRepository: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'******************************************************************************
' FUNCIÓN: ResetMock
' DESCRIPCIÓN: Limpia el mock configurado, volviendo al comportamiento normal
'******************************************************************************
Public Sub ResetMock()
    On Error GoTo ErrorHandler
    Set m_MockRepository = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modRepositoryFactory.ResetMock: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub








