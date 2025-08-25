VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolicitudService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' INTERFAZ: ISolicitudService
' Arquitectura: Contrato Limpio para LÃ³gica de Negocio de Solicitudes
' Version: 2.0 - ReconstrucciÃ³n Total
' ============================================================================
' Define el contrato para operaciones de negocio relacionadas con solicitudes.
' Implementa el patrÃ³n Service con delegaciÃ³n al repositorio.
' ============================================================================

' Crea una nueva solicitud con estado inicial "Borrador"
' @param idExpediente: ID del expediente al que pertenece la solicitud
' @param tipo: Tipo de solicitud (ej: "PC", "RAC", etc.)
' @return: Objeto T_Solicitud creado y guardado
Public Function CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String) As T_Solicitud
End Function

' Guarda una solicitud existente
' @param solicitud: Objeto T_Solicitud a guardar
' @return: True si se guardÃ³ correctamente, False en caso contrario
Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Boolean
End Function











