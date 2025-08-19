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





