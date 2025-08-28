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
' Arquitectura: Contrato Limpio para Lógica de Negocio de Solicitudes
' Version: 2.0 - Reconstrucción Total
' ============================================================================
' Define el contrato para operaciones de negocio relacionadas con solicitudes.
' Implementa el patrón Service con delegación al repositorio.
' ============================================================================

' Crea una nueva solicitud con estado inicial "Borrador"
' @param idExpediente: ID del expediente al que pertenece la solicitud
' @param tipo: Tipo de solicitud (ej: "PC", "RAC", etc.)
' @return: Objeto ESolicitud creado y guardado
Public Function CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String) As ESolicitud
End Function

' Guarda una solicitud existente
' @param solicitud: Objeto ESolicitud a guardar
' @return: True si se guardó correctamente, False en caso contrario
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Boolean
End Function













