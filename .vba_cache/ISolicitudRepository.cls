VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' INTERFAZ: ISolicitudRepository
' Arquitectura: Contrato Limpio para Acceso a Datos de Solicitudes
' Version: 2.0 - Reconstrucción Total
' ============================================================================
' Define el contrato para operaciones de persistencia de solicitudes.
' Implementa el patrón Repository con conexión explícita al Backend.
' ============================================================================

' Obtiene una solicitud por su ID desde el Backend
' @param idSolicitud: ID único de la solicitud
' @return: Objeto T_Solicitud o Nothing si no existe
Public Function GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
End Function

' Guarda una solicitud en el Backend (INSERT o UPDATE)
' @param solicitud: Objeto T_Solicitud a persistir
' @return: ID de la solicitud guardada
Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
End Function






