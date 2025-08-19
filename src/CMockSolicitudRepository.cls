VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockSolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

Implements ISolicitudRepository

Private m_storage As Collection

Private Sub Class_Initialize()
    Set m_storage = New Collection
End Sub

Private Function ISolicitudRepository_LoadPC(ByVal SolicitudID As Long) As CSolicitudPC
    On Error Resume Next
    Set ISolicitudRepository_LoadPC = m_storage(CStr(SolicitudID))
End Function

Private Function ISolicitudRepository_SavePC(solicitudPC As CSolicitudPC) As Long
    Dim newId As Long
    newId = m_storage.Count + 1
    solicitudPC.idSolicitud = newId
    m_storage.Add solicitudPC, CStr(newId)
    ISolicitudRepository_SavePC = newId
End Function

Private Function ISolicitudRepository_ObtenerPorId(ByVal idSolicitud As Long) As T_Solicitud
    On Error Resume Next
    Set ISolicitudRepository_ObtenerPorId = m_storage(CStr(idSolicitud))
End Function





