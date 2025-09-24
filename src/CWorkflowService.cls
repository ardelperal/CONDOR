<<<<<<< HEAD
﻿Attribute VB_GlobalNameSpace = False
=======
﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowService"
Attribute VB_GlobalNameSpace = False
>>>>>>> main
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Implements IWorkflowService

Private m_Repo As IWorkflowRepository
Private m_Logger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal repo As IWorkflowRepository, ByVal logger As IOperationLogger, ByVal ErrorHandler As IErrorHandlerService)
    Set m_Repo = repo
    Set m_Logger = logger
    Set m_ErrorHandler = ErrorHandler
End Sub

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim isValid As Boolean
    
    ' REGLA 1: (MÁXIMA PRIORIDAD) Ninguna transición puede salir de un estado final.
    If IWorkflowService_IsEstadoFinal(estadoOrigen) Then
        isValid = False
    ' REGLA 2: (PRIORIDAD MEDIA) Administrador y Calidad pueden realizar cualquier transición NO-FINAL.
    ElseIf usuarioRol = "Administrador" Or usuarioRol = "Calidad" Then
        isValid = True
    ' REGLA 3: (PRIORIDAD BAJA) Para otros roles, consultar al repositorio.
    Else
        isValid = m_Repo.IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino, usuarioRol)
    End If
    
    ' Registrar la operación de validación
    Dim logEntry As EOperationLog
    Set logEntry = New EOperationLog
    With logEntry
        .tipoOperacion = "VALIDATE_TRANSITION"
        .idEntidadAfectada = SolicitudID
        .usuario = "" ' No disponible en este contexto, dejar vacío
        .resultado = IIf(isValid, "Success", "Failure")
        .descripcion = "Validación de transición: " & estadoOrigen & " -> " & estadoDestino & " para rol " & usuarioRol
    End With
    m_Logger.LogOperation logEntry
    
    IWorkflowService_ValidateTransition = isValid
    
Cleanup:
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowService.ValidateTransition: " & Err.Description, "CWorkflowService"
    IWorkflowService_ValidateTransition = False
    GoTo Cleanup
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
    Dim IdEstado As Long
    
    ' Mapeo de nombres de estados a IDs para el nuevo flujo
    Select Case estadoActual
        Case "Registrado": IdEstado = 1
        Case "Desarrollo": IdEstado = 2
        Case "Modificación": IdEstado = 3
        Case "Validación": IdEstado = 4
        Case "Revisión": IdEstado = 5
        Case "Formalización": IdEstado = 6
        Case "Aprobada": IdEstado = 7
        Case Else
            ' Estado no reconocido
            m_ErrorHandler.LogError 0, "Estado no reconocido: " & estadoActual, "CWorkflowService"
            Set IWorkflowService_GetNextStates = dict
            Exit Function
    End Select
    
    Set dict = m_Repo.GetNextStates(IdEstado, usuarioRol)
    Set IWorkflowService_GetNextStates = dict
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowService.GetNextStates: " & Err.Description, "CWorkflowService"
    Set IWorkflowService_GetNextStates = New Scripting.Dictionary
End Function

Private Function IWorkflowService_IsEstadoFinal(ByVal estadoActual As String) As Boolean
    ' Implementación para verificar si un estado es final en el nuevo flujo
    IWorkflowService_IsEstadoFinal = (estadoActual = "Aprobada")
<<<<<<< HEAD
End Function
=======
End Function




















>>>>>>> main
