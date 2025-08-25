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
' Version: 2.0 - ReconstrucciÃ³n Total
' ============================================================================
' Define el contrato para operaciones de persistencia de solicitudes.
' Implementa el patrÃ³n Repository con conexiÃ³n explÃ­cita al Backend.
' ============================================================================

' Obtiene una solicitud por su ID desde el Backend
' @param idSolicitud: ID Ãºnico de la solicitud
' @return: Objeto T_Solicitud o Nothing si no existe
Public Function GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
End Function

' Guarda una solicitud en el Backend (INSERT o UPDATE)
' @param solicitud: Objeto T_Solicitud a persistir
' @return: ID de la solicitud guardada
Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
End Function

' Ejecuta una consulta SQL con parÃ¡metros opcionales
' @param sql: Consulta SQL a ejecutar
' @param params: ColecciÃ³n opcional de parÃ¡metros QueryParameter
' @return: Recordset con los resultados
Public Function ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.Recordset
End Function






