#If DEV_MODE Then
Option Compare Database
Option Explicit
Implements ISolicitudRepository

Private m_storage As Collection

Private Sub Class_Initialize()
    Set m_storage = New Collection
End Sub

Private Function ISolicitudRepository_LoadPC(solicitudId As Long) As CSolicitudPC
    On Error Resume Next
    Set ISolicitudRepository_LoadPC = m_storage(CStr(solicitudId))
End Function

Private Function ISolicitudRepository_SavePC(solicitudPC As CSolicitudPC) As Long
    Dim newId As Long
    newId = m_storage.Count + 1
    solicitudPC.IdSolicitud = newId
    m_storage.Add solicitudPC, CStr(newId)
    ISolicitudRepository_SavePC = newId
End Function
#End If