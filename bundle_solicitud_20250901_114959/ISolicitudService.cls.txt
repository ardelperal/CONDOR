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
' Version: 3.0 - Refactorización Total
' ============================================================================
' Define el contrato para operaciones de negocio relacionadas con solicitudes.
' Implementa el patrón Service con delegación al repositorio.
' ============================================================================

' Crea una nueva solicitud con estado inicial "Borrador"
' @param paraExpediente: Objeto EExpediente al que pertenece la solicitud
' @return: Objeto ESolicitud creado y guardado
Public Function CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
End Function

' Guarda una solicitud existente
' @param solicitud: Objeto ESolicitud a guardar
' @return: ID de la solicitud guardada (Long), 0 en caso de error
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
End Function

' Obtiene una solicitud por su ID
' @param solicitudId: ID de la solicitud a obtener
' @return: Objeto ESolicitud encontrado
Public Function ObtenerSolicitudPorId(ByVal solicitudId As Long) As ESolicitud
End Function













