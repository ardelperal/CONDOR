Option Compare Database
Option Explicit
Implements ISolicitudRepository

Private m_storage As Collection
Private m_mockRecordset As DAO.Recordset

' Propiedades públicas para verificación en pruebas unitarias
Public GuardarSolicitudCalled As Boolean
Public LastSavedSolicitud As T_Solicitud
Public GuardarSolicitudReturnValue As Long

Private Sub Class_Initialize()
    Set m_storage = New Collection
    GuardarSolicitudCalled = False
    GuardarSolicitudReturnValue = 1 ' Valor por defecto
End Sub

' Método público para inyectar recordset mock en las pruebas
Public Sub SetMockRecordset(recordset As DAO.Recordset)
    Set m_mockRecordset = recordset
End Sub

' Método público para obtener el recordset mock (para uso en AuthService)
Public Function GetMockRecordset() As DAO.Recordset
    Set GetMockRecordset = m_mockRecordset
End Function





' Implementación del método GuardarSolicitud faltante
Private Function ISolicitudRepository_GuardarSolicitud(ByVal solicitud As T_Solicitud) As Long
    ' Mock: registra la llamada y retorna el valor configurado
    GuardarSolicitudCalled = True
    Set LastSavedSolicitud = solicitud
    
    ' Almacenar en la colección interna para compatibilidad
    m_storage.Add solicitud, CStr(GuardarSolicitudReturnValue)
    
    ' Retornar el valor configurado para las pruebas
    ISolicitudRepository_GuardarSolicitud = GuardarSolicitudReturnValue
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



' Implementación del método ExecuteQuery
Private Function ISolicitudRepository_ExecuteQuery(ByVal sql As String) As DAO.Recordset
    ' Mock: devuelve el recordset configurado previamente
    Set ISolicitudRepository_ExecuteQuery = m_mockRecordset
End Function






