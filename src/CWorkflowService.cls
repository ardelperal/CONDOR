VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IWorkflowService

Private m_Repo As IWorkflowRepository
Private m_Logger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal repo As IWorkflowRepository, ByVal logger As IOperationLogger, ByVal errorHandler As IErrorHandlerService)
    Set m_Repo = repo
    Set m_Logger = logger
    Set m_ErrorHandler = errorHandler
End Sub

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    ' Lógica de negocio mejorada para el nuevo flujo de 7 estados
    On Error GoTo ErrorHandler
    
    ' Validaciones de negocio específicas del nuevo flujo
    
    ' 1. No se puede transicionar desde un estado final (Aprobada)
    If estadoOrigen = "Aprobada" Then
        m_Logger.LogOperation "ValidateTransition", SolicitudID, "Intento de transición desde estado final: " & estadoOrigen
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' 2. Validar que la transición existe en el repositorio
    Dim isValidRepo As Boolean
    isValidRepo = m_Repo.IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino)
    
    If Not isValidRepo Then
        m_Logger.LogOperation "ValidateTransition", SolicitudID, "Transición no válida: " & estadoOrigen & " -> " & estadoDestino
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' 3. Log de transición válida
    m_Logger.LogOperation "ValidateTransition", SolicitudID, "Transición válida: " & estadoOrigen & " -> " & estadoDestino & " (Rol: " & usuarioRol & ")"
    
    IWorkflowService_ValidateTransition = True
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowService.ValidateTransition: " & Err.Description, "CWorkflowService"
    IWorkflowService_ValidateTransition = False
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Object
    ' Lógica mejorada para obtener siguientes estados en el flujo de 7 estados
    On Error GoTo ErrorHandler
    
    Dim dict As New Scripting.Dictionary
    
    ' Verificar si es estado final (Aprobada no tiene transiciones de salida)
    If estadoActual = "Aprobada" Then
        Set IWorkflowService_GetNextStates = dict ' Retornar diccionario vacío
        Exit Function
    End If
    
    ' Convertir estadoActual String a Long para el repositorio
    Dim idEstado As Long
    
    ' Mapeo de nombres de estados a IDs para el nuevo flujo
    Select Case estadoActual
        Case "Registrado": idEstado = 1
        Case "Desarrollo": idEstado = 2
        Case "Modificación": idEstado = 3
        Case "Validación": idEstado = 4
        Case "Revisión": idEstado = 5
        Case "Formalización": idEstado = 6
        Case "Aprobada": idEstado = 7
        Case Else
            ' Estado no reconocido
            m_ErrorHandler.LogError 0, "Estado no reconocido: " & estadoActual, "CWorkflowService"
            Set IWorkflowService_GetNextStates = dict
            Exit Function
    End Select
    
    Set dict = m_Repo.GetNextStates(idEstado, usuarioRol)
    Set IWorkflowService_GetNextStates = dict
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowService.GetNextStates: " & Err.Description, "CWorkflowService"
    Set IWorkflowService_GetNextStates = New Scripting.Dictionary
End Function

Private Function IWorkflowService_IsEstadoFinal(ByVal estadoActual As String) As Boolean
    ' Implementación para verificar si un estado es final en el nuevo flujo
    IWorkflowService_IsEstadoFinal = (estadoActual = "Aprobada")
End Function


















