Option Compare Database
Option Explicit
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
    
    ' Obtener la instancia del logger de operaciones
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Usar el repositorio real con inyección de dependencias
    Dim repositoryInstance As New CSolicitudRepository
    
    ' Inyectar la dependencia del logger
    repositoryInstance.Initialize operationLogger
    
    Set CreateSolicitudRepository = repositoryInstance
    
    Exit Function
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "modRepositoryFactory.CreateSolicitudRepository")
    Set CreateSolicitudRepository = Nothing
End Function

'******************************************************************************
' GESTIÓN DE MOCKS PARA PRUEBAS
'******************************************************************************

' Variable privada para almacenar el mock
Private m_MockRepository As ISolicitudRepository

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






