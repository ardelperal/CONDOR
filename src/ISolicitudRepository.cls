Option Compare Database
Option Explicit

' Interfaz para el repositorio de solicitudes
' Define los métodos de persistencia

Public Function SavePC(solicitudPC As CSolicitudPC) As Long
End Function

Public Function LoadPC(ByVal solicitudId As Long) As CSolicitudPC
End Function

Public Function ObtenerPorId(ByVal idSolicitud As Long) As T_Solicitud
End Function