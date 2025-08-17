VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' INTERFAZ: IWorkflowService
' DESCRIPCIÓN: Define el contrato para el servicio de workflow y estados
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTAS: Lección 1 - Implementación de interfaces en VBA
'        Esta interfaz define todos los métodos que debe implementar
'        cualquier servicio de workflow en el sistema CONDOR
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' MÉTODOS DE VALIDACIÓN DE TRANSICIONES
'******************************************************************************

'**
' Valida si una transición de estado es permitida
' @param solicitudId: ID de la solicitud
' @param estadoOrigen: Estado actual de la solicitud
' @param estadoDestino: Estado al que se quiere transicionar
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @param usuarioRol: Rol del usuario que intenta la transición
' @return Boolean: True si la transición es válida, False en caso contrario
'**
Public Function ValidateTransition(ByVal solicitudId As Long, _
                                  ByVal estadoOrigen As String, _
                                  ByVal estadoDestino As String, _
                                  ByVal tipoSolicitud As String, _
                                  ByVal usuarioRol As String) As Boolean
End Function

'******************************************************************************
' MÉTODOS DE OBTENCIÓN DE ESTADOS
'******************************************************************************

'**
' Obtiene todos los estados disponibles para un tipo de solicitud
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @return Collection: Colección de estados disponibles
'**
Public Function GetAvailableStates(ByVal tipoSolicitud As String) As Collection
End Function

'**
' Obtiene los estados siguientes válidos desde un estado actual
' @param estadoActual: Estado actual de la solicitud
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @param usuarioRol: Rol del usuario para filtrar transiciones permitidas
' @return Collection: Colección de estados siguientes válidos
'**
Public Function GetNextStates(ByVal estadoActual As String, _
                             ByVal tipoSolicitud As String, _
                             ByVal usuarioRol As String) As Collection
End Function

'**
' Obtiene el estado inicial para un tipo de solicitud
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @return String: Código del estado inicial
'**
Public Function GetInitialState(ByVal tipoSolicitud As String) As String
End Function

'**
' Verifica si un estado es final para un tipo de solicitud
' @param estadoCodigo: Código del estado a verificar
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @return Boolean: True si es estado final, False en caso contrario
'**
Public Function IsStateFinal(ByVal estadoCodigo As String, _
                            ByVal tipoSolicitud As String) As Boolean
End Function

'******************************************************************************
' MÉTODOS DE HISTORIAL DE ESTADOS
'******************************************************************************

'**
' Registra un cambio de estado en el historial
' @param solicitudId: ID de la solicitud
' @param estadoAnterior: Estado anterior de la solicitud
' @param estadoNuevo: Nuevo estado de la solicitud
' @param usuario: Usuario que realizó el cambio
' @param comentarios: Comentarios opcionales sobre el cambio
' @return Boolean: True si se registró correctamente, False en caso contrario
'**
Public Function RecordStateChange(ByVal solicitudId As Long, _
                                 ByVal estadoAnterior As String, _
                                 ByVal estadoNuevo As String, _
                                 ByVal usuario As String, _
                                 ByVal comentarios As String) As Boolean
End Function

'**
' Obtiene el historial de cambios de estado de una solicitud
' @param solicitudId: ID de la solicitud
' @return Collection: Colección con el historial de cambios
'**
Public Function GetStateHistory(ByVal solicitudId As Long) As Collection
End Function

'******************************************************************************
' MÉTODOS DE VALIDACIÓN DE PERMISOS
'******************************************************************************

'**
' Verifica si un usuario tiene permisos para realizar una transición específica
' @param usuarioRol: Rol del usuario
' @param estadoOrigen: Estado origen de la transición
' @param estadoDestino: Estado destino de la transición
' @param tipoSolicitud: Tipo de solicitud
' @return Boolean: True si tiene permisos, False en caso contrario
'**
Public Function HasTransitionPermission(ByVal usuarioRol As String, _
                                       ByVal estadoOrigen As String, _
                                       ByVal estadoDestino As String, _
                                       ByVal tipoSolicitud As String) As Boolean
End Function

'**
' Verifica si una transición requiere aprobación
' @param estadoOrigen: Estado origen de la transición
' @param estadoDestino: Estado destino de la transición
' @param tipoSolicitud: Tipo de solicitud
' @return Boolean: True si requiere aprobación, False en caso contrario
'**
Public Function RequiresApproval(ByVal estadoOrigen As String, _
                                ByVal estadoDestino As String, _
                                ByVal tipoSolicitud As String) As Boolean
End Function