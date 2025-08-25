Attribute VB_Name = "modWorkflowRepositoryFactory"
'==============================================================================
' MÃ³dulo: modWorkflowRepositoryFactory
' PropÃ³sito: Factory para crear instancias de IWorkflowRepository
' Autor: CONDOR-Expert
' Fecha: 2024
'==============================================================================

Option Compare Database
Option Explicit

' Variable privada para almacenar el mock durante las pruebas
Private m_MockRepository As IWorkflowRepository

'==============================================================================
' FUNCIONES PÃšBLICAS
'==============================================================================

'''
' Crea una instancia de IWorkflowRepository
' Durante las pruebas, devuelve el mock si estÃ¡ configurado
' En producciÃ³n, devuelve la implementaciÃ³n real CWorkflowRepository
' @return IWorkflowRepository: Instancia del repositorio de workflow
'''
Public Function CreateWorkflowRepository() As IWorkflowRepository
    On Error GoTo ErrorHandler
    
    ' Si hay un mock configurado (para pruebas), devolverlo
    If Not m_MockRepository Is Nothing Then
        Set CreateWorkflowRepository = m_MockRepository
        Exit Function
    End If
    
    ' En caso contrario, crear la implementaciÃ³n real
    Dim repository As CWorkflowRepository
    Set repository = New CWorkflowRepository
    
    ' Obtener instancia de configuraciÃ³n usando el nuevo factory
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    repository.Initialize configService
    
    Set CreateWorkflowRepository = repository
    
    Exit Function
    
ErrorHandler:
    Set CreateWorkflowRepository = Nothing
    Debug.Print "Error en CreateWorkflowRepository: " & Err.Number & " - " & Err.Description
End Function

'==============================================================================
' FUNCIONES PARA TESTING
'==============================================================================

'''
' Configura un mock repository para las pruebas
' @param mockRepo: Instancia del mock a usar en lugar de la implementaciÃ³n real
'''
Public Sub SetMockRepository(ByVal mockRepo As IWorkflowRepository)
    Set m_MockRepository = mockRepo
End Sub

'''
' Resetea el mock repository, volviendo a usar la implementaciÃ³n real
'''
Public Sub ResetMock()
    Set m_MockRepository = Nothing
End Sub

'''
' Verifica si hay un mock configurado
' @return Boolean: True si hay un mock activo, False en caso contrario
'''
Public Function HasMock() As Boolean
    HasMock = Not (m_MockRepository Is Nothing)
End Function
