Attribute VB_Name = "modWorkflowRepositoryFactory"
Option Compare Database
Option Explicit

'==============================================================================
' Módulo: modWorkflowRepositoryFactory
' Propósito: Factory para crear instancias de IWorkflowRepository
' Autor: CONDOR-Expert
' Fecha: 2024
'==============================================================================


' Variable privada para almacenar el mock durante las pruebas
Private m_MockRepository As IWorkflowRepository

'==============================================================================
' FUNCIONES PÚBLICAS
'==============================================================================

'''
' Crea una instancia de IWorkflowRepository
' Durante las pruebas, devuelve el mock si está configurado
' En producción, devuelve la implementación real CWorkflowRepository
' @param errorHandler: Servicio de manejo de errores
' @return IWorkflowRepository: Instancia del repositorio de workflow
'''
Public Function CreateWorkflowRepository(ByVal errorHandler As IErrorHandlerService) As IWorkflowRepository
    On Error GoTo ErrorHandler
    
    ' Si hay un mock configurado (para pruebas), devolverlo
    If Not m_MockRepository Is Nothing Then
        Set CreateWorkflowRepository = m_MockRepository
        Exit Function
    End If
    
    ' En caso contrario, crear la implementación real
    Dim repository As CWorkflowRepository
    Set repository = New CWorkflowRepository
    
    ' Obtener instancia de configuración usando el nuevo factory
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService(errorHandler)
    repository.Initialize configService
    
    Set CreateWorkflowRepository = repository
    
    Exit Function
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modWorkflowRepositoryFactory.CreateWorkflowRepository"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'==============================================================================
' FUNCIONES PARA TESTING
'==============================================================================

'''
' Configura un mock repository para las pruebas
' @param mockRepo: Instancia del mock a usar en lugar de la implementación real
'''
Public Sub SetMockRepository(ByVal mockRepo As IWorkflowRepository)
    On Error GoTo ErrorHandler
    Set m_MockRepository = mockRepo
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modWorkflowRepositoryFactory.SetMockRepository: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'''
' Resetea el mock repository, volviendo a usar la implementación real
'''
Public Sub ResetMock()
    On Error GoTo ErrorHandler
    Set m_MockRepository = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modWorkflowRepositoryFactory.ResetMock: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'''
' Verifica si hay un mock configurado
' @return Boolean: True si hay un mock activo, False en caso contrario
'''
Public Function HasMock() As Boolean
    On Error GoTo ErrorHandler
    HasMock = Not (m_MockRepository Is Nothing)
    Exit Function
ErrorHandler:
    Debug.Print "Error en modWorkflowRepositoryFactory.HasMock: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


