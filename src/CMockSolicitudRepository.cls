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

'******************************************************************************
' CLASE: CMockSolicitudRepository
' DESCRIPCIÓN: Implementación Mock del repositorio de solicitudes para pruebas unitarias.
' AUTOR: Sistema CONDOR
' FECHA: 2025-08-27
'******************************************************************************

Implements ISolicitudRepository

' Variables para tracking de llamadas y configuración de retorno
Private m_GetSolicitudById_WasCalled As Boolean
Private m_SaveSolicitud_WasCalled As Boolean
Private m_ExecuteQuery_WasCalled As Boolean

Private m_LastIdSolicitud As Long
Private m_LastSolicitud As T_Solicitud
Private m_LastSql As String
Private m_LastParams As Collection

Private m_GetSolicitudByIdReturnValue As T_Solicitud
Private m_SaveSolicitudReturnValue As Long
Private m_ExecuteQueryReturnValue As DAO.recordset

Private m_GetSolicitudByIdCallCount As Long
Private m_SaveSolicitudCallCount As Long
Private m_ExecuteQueryCallCount As Long

'--- PROPIEDADES PARA VERIFICACIÓN DE LLAMADAS ---

Public Property Get GetSolicitudById_WasCalled() As Boolean
    GetSolicitudById_WasCalled = m_GetSolicitudById_WasCalled
End Property

Public Property Get SaveSolicitud_WasCalled() As Boolean
    SaveSolicitud_WasCalled = m_SaveSolicitud_WasCalled
End Property

Public Property Get ExecuteQuery_WasCalled() As Boolean
    ExecuteQuery_WasCalled = m_ExecuteQuery_WasCalled
End Property

Public Property Get LastIdSolicitud() As Long
    LastIdSolicitud = m_LastIdSolicitud
End Property

Public Property Get LastSolicitud() As T_Solicitud
    Set LastSolicitud = m_LastSolicitud
End Property

Public Property Get LastSql() As String
    LastSql = m_LastSql
End Property

Public Property Get LastParams() As Collection
    Set LastParams = m_LastParams
End Property

Public Property Get GetSolicitudByIdCallCount() As Long
    GetSolicitudByIdCallCount = m_GetSolicitudByIdCallCount
End Property

Public Property Get SaveSolicitudCallCount() As Long
    SaveSolicitudCallCount = m_SaveSolicitudCallCount
End Property

Public Property Get ExecuteQueryCallCount() As Long
    ExecuteQueryCallCount = m_ExecuteQueryCallCount
End Property

'--- MÉTODOS PARA CONFIGURAR VALORES DE RETORNO ---

Public Sub SetGetSolicitudByIdReturnValue(ByVal returnValue As T_Solicitud)
    Set m_GetSolicitudByIdReturnValue = returnValue
End Sub

Public Sub SetSaveSolicitudReturnValue(ByVal returnValue As Long)
    m_SaveSolicitudReturnValue = returnValue
End Sub

Public Sub SetExecuteQueryReturnValue(ByVal returnValue As DAO.recordset)
    Set m_ExecuteQueryReturnValue = returnValue
End Sub

'--- MÉTODO RESET ---

Public Sub Reset()
    m_GetSolicitudById_WasCalled = False
    m_SaveSolicitud_WasCalled = False
    m_ExecuteQuery_WasCalled = False
    
    m_LastIdSolicitud = 0
    Set m_LastSolicitud = Nothing
    m_LastSql = ""
    Set m_LastParams = Nothing
    
    Set m_GetSolicitudByIdReturnValue = Nothing
    m_SaveSolicitudReturnValue = 0
    Set m_ExecuteQueryReturnValue = Nothing
    
    m_GetSolicitudByIdCallCount = 0
    m_SaveSolicitudCallCount = 0
    m_ExecuteQueryCallCount = 0
End Sub

'--- IMPLEMENTACIÓN DE LA INTERFAZ ISolicitudRepository ---

Private Function ISolicitudRepository_GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    m_GetSolicitudById_WasCalled = True
    m_LastIdSolicitud = idSolicitud
    m_GetSolicitudByIdCallCount = m_GetSolicitudByIdCallCount + 1
    
    Set ISolicitudRepository_GetSolicitudById = m_GetSolicitudByIdReturnValue
End Function

Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    m_SaveSolicitud_WasCalled = True
    Set m_LastSolicitud = solicitud
    m_SaveSolicitudCallCount = m_SaveSolicitudCallCount + 1
    
    ISolicitudRepository_SaveSolicitud = m_SaveSolicitudReturnValue
End Function

Private Function ISolicitudRepository_ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.recordset
    m_ExecuteQuery_WasCalled = True
    m_LastSql = sql
    Set m_LastParams = params
    m_ExecuteQueryCallCount = m_ExecuteQueryCallCount + 1
    
    Set ISolicitudRepository_ExecuteQuery = m_ExecuteQueryReturnValue
End Function

'--- MÉTODOS PÚBLICOS DE CONVENIENCIA ---

Public Function GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    Set GetSolicitudById = ISolicitudRepository_GetSolicitudById(idSolicitud)
End Function

Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function

Public Function ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.recordset
    Set ExecuteQuery = ISolicitudRepository_ExecuteQuery(sql, params)
End Function