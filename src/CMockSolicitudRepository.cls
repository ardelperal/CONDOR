Option Compare Database
Option Explicit
Implements ISolicitudRepository

Private m_storage As Collection

Private Sub Class_Initialize()
    Set m_storage = New Collection
End Sub





' Implementación del método GuardarSolicitud faltante
Private Function ISolicitudRepository_GuardarSolicitud(ByVal solicitud As T_Solicitud) As Long
    ' Mock: registra la llamada y retorna un ID simulado
    Dim newId As Long
    newId = m_storage.Count + 1
    m_storage.Add solicitud, CStr(newId)
    ISolicitudRepository_GuardarSolicitud = newId
End Function

' Implementación del método ObtenerSolicitudPorId
Private Function ISolicitudRepository_ObtenerSolicitudPorId(ByVal idSolicitud As Long) As T_Solicitud
    On Error Resume Next
    Set ISolicitudRepository_ObtenerSolicitudPorId = m_storage(CStr(idSolicitud))
End Function

' Implementación del método ObtenerSolicitudPorCodigo
Private Function ISolicitudRepository_ObtenerSolicitudPorCodigo(ByVal codigoSolicitud As String) As T_Solicitud
    ' Mock: retorna Nothing (no implementado)
    Set ISolicitudRepository_ObtenerSolicitudPorCodigo = Nothing
End Function

' Implementación del método EliminarSolicitud
Private Function ISolicitudRepository_EliminarSolicitud(ByVal idSolicitud As Long) As Boolean
    ' Mock: simula eliminación exitosa
    ISolicitudRepository_EliminarSolicitud = True
End Function

' Implementación del método ObtenerSolicitudesPorExpediente
Private Function ISolicitudRepository_ObtenerSolicitudesPorExpediente(ByVal idExpediente As String) As Collection
    ' Mock: retorna colección vacía
    Set ISolicitudRepository_ObtenerSolicitudesPorExpediente = New Collection
End Function

' Implementación del método ExisteSolicitudConCodigo
Private Function ISolicitudRepository_ExisteSolicitudConCodigo(ByVal codigoSolicitud As String) As Boolean
    ' Mock: retorna False por defecto
    ISolicitudRepository_ExisteSolicitudConCodigo = False
End Function

' Implementación del método Load faltante
Private Function ISolicitudRepository_Load(ByVal ID As Long) As Object
    ' Mock: no hace nada, solo cumple el contrato
    Set ISolicitudRepository_Load = Nothing
End Function

' Implementación del método Save faltante
Private Function ISolicitudRepository_Save(ByVal solicitud As ISolicitud) As Long
    ' Mock: simula guardado exitoso y retorna ID simulado
    ISolicitudRepository_Save = m_storage.Count + 1
End Function






