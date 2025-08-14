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

Private Function ISolicitudService_CreateNuevaSolicitud(ByVal tipoSolicitud As String) As ISolicitud
    ' Simula la creación de una nueva solicitud.
    ' Devolvemos una instancia de la clase de solicitud concreta para PC.
    Set ISolicitudService_CreateNuevaSolicitud = New CSolicitudPC
End Function

Private Function ISolicitudService_GetSolicitudPorID(ByVal id As Long) As ISolicitud
    ' Simula la obtención de una solicitud.
    ' Devolvemos un objeto para que la prueba no falle.
    Set ISolicitudService_GetSolicitudPorID = New CSolicitudPC
End Function

Private Function ISolicitudService_SaveSolicitud(ByVal solicitud As ISolicitud) As Boolean
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

Private Function ISolicitudService_DeleteSolicitud(ByVal id As Long) As Boolean
    ' Simula el borrado, siempre devuelve éxito.
    ISolicitudService_DeleteSolicitud = True
End Function

Private Function ISolicitudService_UpdateEstadoSolicitud(ByVal id As Long, ByVal nuevoEstado As String) As Boolean
    ' Simula la actualización de estado, siempre devuelve éxito.
    ISolicitudService_UpdateEstadoSolicitud = True
End Function
