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
' DESCRIPCI├ôN: Define el contrato para el servicio de workflow y estados
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTAS: Lecci├│n 1 - Implementaci├│n de interfaces en VBA
'        Esta interfaz define todos los m├®todos que debe implementar
'        cualquier servicio de workflow en el sistema CONDOR
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' M├ëTODOS DE VALIDACI├ôN DE TRANSICIONES
'******************************************************************************

'**
' Valida si una transici├│n de estado es permitida
' @param solicitudId: ID de la solicitud
' @param estadoOrigen: Estado actual de la solicitud
' @param estadoDestino: Estado al que se quiere transicionar
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @param usuarioRol: Rol del usuario que intenta la transici├│n
' @return Boolean: True si la transici├│n es v├ílida, False en caso contrario
'**
Public Function ValidateTransition(ByVal solicitudId As Long, _
                                  ByVal estadoOrigen As String, _
                                  ByVal estadoDestino As String, _
                                  ByVal tipoSolicitud As String, _
                                  ByVal usuarioRol As String) As Boolean
End Function

'******************************************************************************
' M├ëTODOS DE OBTENCI├ôN DE ESTADOS
'******************************************************************************

'**
' Obtiene todos los estados disponibles para un tipo de solicitud
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @return Collection: Colecci├│n de estados disponibles
'**
Public Function GetAvailableStates(ByVal tipoSolicitud As String) As Collection
End Function

'**
' Obtiene los estados siguientes v├ílidos desde un estado actual
' @param estadoActual: Estado actual de la solicitud
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @param usuarioRol: Rol del usuario para filtrar transiciones permitidas
' @return Collection: Colecci├│n de estados siguientes v├ílidos
'**
Public Function GetNextStates(ByVal estadoActual As String, _
                             ByVal tipoSolicitud As String, _
                             ByVal usuarioRol As String) As Collection
End Function

'**
' Obtiene el estado inicial para un tipo de solicitud
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @return String: C├│digo del estado inicial
'**
Public Function GetInitialState(ByVal tipoSolicitud As String) As String
End Function

'**
' Verifica si un estado es final para un tipo de solicitud
' @param estadoCodigo: C├│digo del estado a verificar
' @param tipoSolicitud: Tipo de solicitud (PC, Software, etc.)
' @return Boolean: True si es estado final, False en caso contrario
'**
Public Function IsStateFinal(ByVal estadoCodigo As String, _
                            ByVal tipoSolicitud As String) As Boolean
End Function

'******************************************************************************
' M├ëTODOS DE HISTORIAL DE ESTADOS
'******************************************************************************

'**
' Registra un cambio de estado en el historial
' @param solicitudId: ID de la solicitud
' @param estadoAnterior: Estado anterior de la solicitud
' @param estadoNuevo: Nuevo estado de la solicitud
' @param usuario: Usuario que realiz├│ el cambio
' @param comentarios: Comentarios opcionales sobre el cambio
' @return Boolean: True si se registr├│ correctamente, False en caso contrario
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
' @return Collection: Colecci├│n con el historial de cambios
'**
Public Function GetStateHistory(ByVal solicitudId As Long) As Collection
End Function

'******************************************************************************
' M├ëTODOS DE VALIDACI├ôN DE PERMISOS
'******************************************************************************

'**
' Verifica si un usuario tiene permisos para realizar una transici├│n espec├¡fica
' @param usuarioRol: Rol del usuario
' @param estadoOrigen: Estado origen de la transici├│n
' @param estadoDestino: Estado destino de la transici├│n
' @param tipoSolicitud: Tipo de solicitud
' @return Boolean: True si tiene permisos, False en caso contrario
'**
Public Function HasTransitionPermission(ByVal usuarioRol As String, _
                                       ByVal estadoOrigen As String, _
                                       ByVal estadoDestino As String, _
                                       ByVal tipoSolicitud As String) As Boolean
End Function

'**
' Verifica si una transici├│n requiere aprobaci├│n
' @param estadoOrigen: Estado origen de la transici├│n
' @param estadoDestino: Estado destino de la transici├│n
' @param tipoSolicitud: Tipo de solicitud
' @return Boolean: True si requiere aprobaci├│n, False en caso contrario
'**
Public Function RequiresApproval(ByVal estadoOrigen As String, _
                                ByVal estadoDestino As String, _
                                ByVal tipoSolicitud As String) As Boolean
End Function
