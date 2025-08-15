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

' Interfaz para el Servicio de Solicitudes
' Define los m?todos principales para manejar solicitudes en el sistema CONDOR
' Version: 1.0
' Fecha: 2025-01-14


' M?todo para crear una nueva solicitud
Public Function CreateNuevaSolicitud(ByVal tipo As String) As ISolicitud
End Function

' M?todo para obtener una solicitud por su ID
Public Function GetSolicitudPorID(ByVal ID As Long) As ISolicitud
End Function

' M?todo para guardar una solicitud
Public Function SaveSolicitud(ByVal solicitud As ISolicitud) As Boolean
End Function

' M?todo para obtener todas las solicitudes
Public Function GetAllSolicitudes() As Collection
End Function

' M?todo para eliminar una solicitud
Public Function DeleteSolicitud(ByVal ID As Long) As Boolean
End Function

' M?todo para actualizar el estado de una solicitud
Public Function UpdateEstadoSolicitud(ByVal ID As Long, ByVal nuevoEstado As String) As Boolean
End Function



