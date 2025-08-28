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
' @param params: Collection de objetos QueryParameter con los parámetros de la consulta
' @return: Objeto ESolicitud o Nothing si no existe
Public Function GetSolicitudById(ByVal params As Collection) As ESolicitud
End Function

' Guarda una solicitud en el Backend (INSERT o UPDATE)
' @param params: Collection de objetos QueryParameter con los datos de la solicitud
' @return: ID de la solicitud guardada
Public Function SaveSolicitud(ByVal params As Collection) As Long
End Function

' Ejecuta una consulta SQL con parámetros
' @param queryName: Nombre de la consulta desde modQueries
' @param params: Collection de objetos QueryParameter
' @return: Recordset con los resultados
Public Function ExecuteQuery(ByVal queryName As String, ByVal params As Collection) As DAO.recordset
End Function








