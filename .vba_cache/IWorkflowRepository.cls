VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'==============================================================================
' Interfaz: IWorkflowRepository
' Propósito: Define el contrato para gestionar las reglas de transición de estados
'            del workflow de solicitudes
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'==============================================================================


'==============================================================================
' MÉTODOS PÚBLICOS
'==============================================================================

' Verifica si una transición de estado es válida según las reglas del workflow
' @param tipoSolicitud: Tipo de solicitud (ej: "PC", "CD_CA")
' @param estadoOrigen: Estado actual de la solicitud
' @param estadoDestino: Estado al que se quiere transicionar
' @return: True si la transición es válida, False en caso contrario
Public Function IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
End Function

' Obtiene todos los estados disponibles para un tipo de solicitud
' @param tipoSolicitud: Tipo de solicitud
' @return: Collection con los códigos de estados disponibles
Public Function GetAvailableStates(ByVal tipoSolicitud As String) As Scripting.Dictionary
End Function

' Obtiene los estados siguientes válidos desde un estado actual
' @param estadoActual: Estado actual de la solicitud
' @param tipoSolicitud: Tipo de solicitud
' @param usuarioRol: Rol del usuario que solicita la transición
' @return: Collection con los códigos de estados siguientes válidos
Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Scripting.Dictionary
End Function

' Obtiene el estado inicial para un tipo de solicitud
' @param tipoSolicitud: Tipo de solicitud
' @return: Código del estado inicial
Public Function GetInitialState(ByVal tipoSolicitud As String) As String
End Function

' Verifica si un estado es final para un tipo de solicitud
' @param estadoCodigo: Código del estado a verificar
' @param tipoSolicitud: Tipo de solicitud
' @return: True si el estado es final, False en caso contrario
Public Function IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
End Function

' Registra un cambio de estado en el historial
' @param SolicitudID: ID de la solicitud
' @param estadoAnterior: Estado anterior
' @param estadoNuevo: Estado nuevo
' @param usuario: Usuario que realiza el cambio
' @param comentarios: Comentarios del cambio
' @return: True si se registró correctamente, False en caso contrario
Public Function RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
End Function

' Obtiene el historial de cambios de estado de una solicitud
' @param SolicitudID: ID de la solicitud
' @return: Collection con el historial de cambios
Public Function GetStateHistory(ByVal SolicitudID As Long) As Scripting.Dictionary
End Function

' Verifica si un usuario tiene permisos para realizar una transición
' @param usuarioRol: Rol del usuario
' @param estadoOrigen: Estado origen de la transición
' @param estadoDestino: Estado destino de la transición
' @param tipoSolicitud: Tipo de solicitud
' @return: True si tiene permisos, False en caso contrario
Public Function HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
End Function

' Verifica si una transición requiere aprobación
' @param estadoOrigen: Estado origen de la transición
' @param estadoDestino: Estado destino de la transición
' @param tipoSolicitud: Tipo de solicitud
' @return: True si requiere aprobación, False en caso contrario
Public Function RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
End Function

Public Function GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
End Function


