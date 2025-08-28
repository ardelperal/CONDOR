Attribute VB_Name = "CMockSolicitudRepository"
Option Compare Database
Option Explicit
Implements ISolicitudRepository

' --- Variables de estado para configurar el comportamiento del mock ---
Private m_GetSolicitudById_Result As E_Solicitud
Private m_SaveSolicitud_Result As Long
Private m_ExecuteQuery_Result As DAO.Recordset

' --- Propiedades de seguimiento para las aserciones ---
Public GetSolicitudById_CallCount As Long
Public SaveSolicitud_CallCount As Long
Public ExecuteQuery_CallCount As Long
Public LastSavedSolicitud As E_Solicitud

' --- Métodos públicos para configurar el mock desde los tests ---
Public Sub ConfigureGetSolicitudById(ByVal returnValue As E_Solicitud)
    Set m_GetSolicitudById_Result = returnValue
End Sub

Public Sub ConfigureSaveSolicitud(ByVal returnValue As Long)
    m_SaveSolicitud_Result = returnValue
End Sub

Public Sub ConfigureExecuteQuery(ByVal returnValue As DAO.Recordset)
    Set m_ExecuteQuery_Result = returnValue
End Sub

' --- Resetea el estado del mock para aislar los tests ---
Public Sub Reset()
    Set m_GetSolicitudById_Result = Nothing
    m_SaveSolicitud_Result = 0
    Set m_ExecuteQuery_Result = Nothing
    GetSolicitudById_CallCount = 0
    SaveSolicitud_CallCount = 0
    ExecuteQuery_CallCount = 0
    Set LastSavedSolicitud = Nothing
End Sub

' --- Implementación de la Interfaz ---
Private Function ISolicitudRepository_GetSolicitudById(ByVal idSolicitud As Long) As E_Solicitud
    GetSolicitudById_CallCount = GetSolicitudById_CallCount + 1
    Set ISolicitudRepository_GetSolicitudById = m_GetSolicitudById_Result
End Function

Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As E_Solicitud) As Long
    SaveSolicitud_CallCount = SaveSolicitud_CallCount + 1
    Set LastSavedSolicitud = solicitud
    ISolicitudRepository_SaveSolicitud = m_SaveSolicitud_Result
End Function

Private Function ISolicitudRepository_ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.Recordset
    ExecuteQuery_CallCount = ExecuteQuery_CallCount + 1
    Set ISolicitudRepository_ExecuteQuery = m_ExecuteQuery_Result
End Function