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

' --- Variables de estado para configurar el comportamiento del mock ---
Private m_GetSolicitudById_Result As ESolicitud
Private m_SaveSolicitud_Result As Long

' --- Propiedades de seguimiento para las aserciones ---
Public GetSolicitudById_CallCount As Long
Public SaveSolicitud_CallCount As Long
Public LastSavedSolicitud As ESolicitud

' --- Métodos públicos para configurar el mock desde los tests ---
Public Sub ConfigureGetSolicitudById(ByVal returnValue As ESolicitud)
    Set m_GetSolicitudById_Result = returnValue
End Sub

Public Sub ConfigureSaveSolicitud(ByVal returnValue As Long)
    m_SaveSolicitud_Result = returnValue
End Sub

' --- Resetea el estado del mock para aislar los tests ---
Public Sub Reset()
    Set m_GetSolicitudById_Result = Nothing
    m_SaveSolicitud_Result = 0
    GetSolicitudById_CallCount = 0
    SaveSolicitud_CallCount = 0
    Set LastSavedSolicitud = Nothing
End Sub

' --- Implementación de la Interfaz ---
Private Function ISolicitudRepository_GetSolicitudById(ByVal idSolicitud As Long) As ESolicitud
    GetSolicitudById_CallCount = GetSolicitudById_CallCount + 1
    Set ISolicitudRepository_GetSolicitudById = m_GetSolicitudById_Result
End Function

Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud_CallCount = SaveSolicitud_CallCount + 1
    Set LastSavedSolicitud = solicitud
    ISolicitudRepository_SaveSolicitud = m_SaveSolicitud_Result
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================

' Método público de conveniencia para GetSolicitudById
Public Function GetSolicitudById(ByVal idSolicitud As Long) As ESolicitud
    Set GetSolicitudById = ISolicitudRepository_GetSolicitudById(idSolicitud)
End Function

' Método público de conveniencia para SaveSolicitud
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function