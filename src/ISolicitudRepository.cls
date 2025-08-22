Option Compare Database
Option Explicit
' Interfaz para el repositorio de solicitudes
' Define los métodos para la persistencia de solicitudes y sus datos asociados

' Guarda una solicitud completa (INSERT o UPDATE según si existe idSolicitud)
Public Function GuardarSolicitud(ByVal solicitud As T_Solicitud) As Long
End Function

' Recupera una solicitud por su ID
Public Function ObtenerSolicitudPorId(ByVal idSolicitud As Long) As T_Solicitud
End Function

' Recupera una solicitud por su código
Public Function ObtenerSolicitudPorCodigo(ByVal codigoSolicitud As String) As T_Solicitud
End Function

' Elimina una solicitud y todos sus datos asociados
Public Function EliminarSolicitud(ByVal idSolicitud As Long) As Boolean
End Function

' Obtiene todas las solicitudes de un expediente
Public Function ObtenerSolicitudesPorExpediente(ByVal idExpediente As String) As Collection
End Function

' Verifica si existe una solicitud con el código dado
Public Function ExisteSolicitudConCodigo(ByVal codigoSolicitud As String) As Boolean
End Function

' Método genérico para cargar una solicitud por ID (Principio de Abstracción de Interfaces)
Public Function Load(ByVal ID As Long) As Object
End Function

' Método genérico para guardar una solicitud (Principio de Abstracción de Interfaces)
Public Function Save(ByVal solicitud As Object) As Long
End Function

' Ejecuta una consulta SQL y devuelve un recordset
Public Function ExecuteQuery(ByVal sql As String) As DAO.Recordset
End Function






