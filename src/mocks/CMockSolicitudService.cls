Attribute VB_Name = "CMockSolicitudService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase Mock para ISolicitudService
' Simula el comportamiento del servicio de solicitudes para pruebas unitarias.
Option Compare Database
Option Explicit

Implements ISolicitudService

Private Function ISolicitudService_CreateNuevaSolicitud(ByVal tipo As String) As ISolicitud
    ' Simula la creación de una nueva solicitud.
    ' Devolvemos una instancia de la clase de solicitud concreta para PC.
    Set ISolicitudService_CreateNuevaSolicitud = New CSolicitudPC
End Function

Private Function ISolicitudService_GetSolicitudPorID(ByVal ID As Long) As ISolicitud
    ' Simula la obtención de una solicitud.
    ' Devolvemos un objeto para que la prueba no falle.
    Set ISolicitudService_GetSolicitudPorID = New CSolicitudPC
End Function

Private Function ISolicitudService_SaveSolicitud(solicitud As ISolicitud) As Boolean
    ' Simula el guardado, siempre devuelve éxito.
    ISolicitudService_SaveSolicitud = True
End Function

Private Function ISolicitudService_GetAllSolicitudes() As Collection
    ' Simula la obtención de todas las solicitudes.
    ' Devuelve una colección con al menos un objeto.
    Dim col As New Collection
    col.Add New CSolicitudPC
    Set ISolicitudService_GetAllSolicitudes = col
End Function

Private Function ISolicitudService_DeleteSolicitud(ByVal ID As Long) As Boolean
    ' Simula el borrado, siempre devuelve éxito.
    ISolicitudService_DeleteSolicitud = True
End Function

Private Function ISolicitudService_UpdateEstadoSolicitud(ByVal ID As Long, ByVal nuevoEstado As String) As Boolean
    ' Simula la actualización de estado, siempre devuelve éxito.
    ISolicitudService_UpdateEstadoSolicitud = True
End Function

' Métodos adicionales requeridos por ISolicitudService
Private Function ISolicitudService_GetSolicitud(ByVal ID As Long) As ISolicitud
    Set ISolicitudService_GetSolicitud = New CSolicitudPC
End Function

Private Function ISolicitudService_CreateSolicitud(ByVal idExpediente As Long, ByVal tipo As String, ByVal descripcion As String, ByVal idUsuario As Long) As Long
    ISolicitudService_CreateSolicitud = 1001 ' ID simulado
End Function

Private Function ISolicitudService_UpdateSolicitud(ByVal ID As Long, ByVal descripcion As String, ByVal estado As String) As Boolean
    ISolicitudService_UpdateSolicitud = True
End Function

Private Function ISolicitudService_ChangeEstado(ByVal ID As Long, ByVal nuevoEstado As String) As Boolean
    ISolicitudService_ChangeEstado = True
End Function

Private Function ISolicitudService_GetSolicitudesByExpediente(ByVal idExpediente As Long) As Collection
    Set ISolicitudService_GetSolicitudesByExpediente = New Collection
End Function

Private Function ISolicitudService_GetSolicitudesByTipo(ByVal tipo As String) As Collection
    Set ISolicitudService_GetSolicitudesByTipo = New Collection
End Function

Private Function ISolicitudService_GetSolicitudesByEstado(ByVal estado As String) As Collection
    Set ISolicitudService_GetSolicitudesByEstado = New Collection
End Function

Private Function ISolicitudService_SearchSolicitudes(ByVal criterio As String) As Collection
    Set ISolicitudService_SearchSolicitudes = New Collection
End Function

Private Function ISolicitudService_ValidateSolicitud(solicitud As ISolicitud) As Boolean
    ISolicitudService_ValidateSolicitud = True
End Function
