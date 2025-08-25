Option Compare Database
Option Explicit

' ============================================================================
' CLASE: CMockSolicitudRepository
' Arquitectura: Mock del Repositorio de Solicitudes para Pruebas Unitarias
' Version: 2.0 - ReconstrucciÃ³n Total
' ============================================================================
' Mock que implementa ISolicitudRepository para aislamiento de pruebas unitarias.
' Permite verificar llamadas y configurar comportamientos de prueba.
' ============================================================================

Implements ISolicitudRepository

' ============================================================================
' PROPIEDADES PÃšBLICAS PARA VERIFICACIÃ“N EN PRUEBAS
' ============================================================================

' Flags de llamadas
Public GetSolicitudByIdCalled As Boolean
Public SaveSolicitudCalled As Boolean
Public ExecuteQueryCalled As Boolean

' ParÃ¡metros capturados
Public LastIdSolicitudRequested As Long
Public LastSavedSolicitud As T_Solicitud
Public LastExecutedSql As String
Public LastExecutedParams As Collection

' Valores de retorno configurables
Public GetSolicitudByIdReturnValue As T_Solicitud
Public SaveSolicitudReturnValue As Long
Public ExecuteQueryReturnValue As DAO.Recordset

' Contadores
Public GetSolicitudByIdCallCount As Long
Public SaveSolicitudCallCount As Long
Public ExecuteQueryCallCount As Long

' ============================================================================
' CONSTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Reset
End Sub

' ============================================================================
' MÃ‰TODOS PÃšBLICOS PARA CONFIGURACIÃ“N DEL MOCK
' ============================================================================

' Reinicia todas las propiedades del mock a sus valores por defecto
Public Sub Reset()
    ' Resetear flags de llamadas
    GetSolicitudByIdCalled = False
    SaveSolicitudCalled = False
    ExecuteQueryCalled = False
    
    ' Resetear parÃ¡metros capturados
    LastIdSolicitudRequested = 0
    Set LastSavedSolicitud = Nothing
    LastExecutedSql = ""
    Set LastExecutedParams = Nothing
    
    ' Resetear valores de retorno
    Set GetSolicitudByIdReturnValue = Nothing
    SaveSolicitudReturnValue = 1 ' Valor por defecto exitoso
    Set ExecuteQueryReturnValue = Nothing
    
    ' Resetear contadores
    GetSolicitudByIdCallCount = 0
    SaveSolicitudCallCount = 0
    ExecuteQueryCallCount = 0
End Sub

' Configura la solicitud que debe devolver GetSolicitudById
Public Sub SetGetSolicitudByIdReturnValue(ByVal solicitud As T_Solicitud)
    Set GetSolicitudByIdReturnValue = solicitud
End Sub

' Configura el ID que debe devolver SaveSolicitud
Public Sub SetSaveSolicitudReturnValue(ByVal returnId As Long)
    SaveSolicitudReturnValue = returnId
End Sub

' Configura el Recordset que debe devolver ExecuteQuery
Public Sub SetExecuteQueryReturnValue(ByVal rs As DAO.Recordset)
    Set ExecuteQueryReturnValue = rs
End Sub

' ============================================================================
' IMPLEMENTACIÃ“N DE ISolicitudRepository
' ============================================================================

' ImplementaciÃ³n mock de GetSolicitudById
Private Function ISolicitudRepository_GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    ' Registrar la llamada
    GetSolicitudByIdCalled = True
    GetSolicitudByIdCallCount = GetSolicitudByIdCallCount + 1
    LastIdSolicitudRequested = idSolicitud
    
    ' Devolver el valor configurado
    If Not GetSolicitudByIdReturnValue Is Nothing Then
        Set ISolicitudRepository_GetSolicitudById = GetSolicitudByIdReturnValue
    Else
        Set ISolicitudRepository_GetSolicitudById = Nothing
    End If
End Function

' ImplementaciÃ³n mock de SaveSolicitud
Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    ' Registrar la llamada
    SaveSolicitudCalled = True
    SaveSolicitudCallCount = SaveSolicitudCallCount + 1
    Set LastSavedSolicitud = solicitud
    
    ' Devolver el valor configurado
    ISolicitudRepository_SaveSolicitud = SaveSolicitudReturnValue
End Function

' ImplementaciÃ³n mock de ExecuteQuery
Private Function ISolicitudRepository_ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.Recordset
    ' Registrar la llamada
    ExecuteQueryCalled = True
    ExecuteQueryCallCount = ExecuteQueryCallCount + 1
    LastExecutedSql = sql
    Set LastExecutedParams = params
    
    ' Devolver el valor configurado
    Set ISolicitudRepository_ExecuteQuery = ExecuteQueryReturnValue
End Function

' ============================================================================
' MÃ‰TODOS PÃšBLICOS DE CONVENIENCIA (DELEGACIÃ“N)
' ============================================================================

' MÃ©todo pÃºblico que delega a la implementaciÃ³n de la interfaz
Public Function GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    Set GetSolicitudById = ISolicitudRepository_GetSolicitudById(idSolicitud)
End Function

' MÃ©todo pÃºblico que delega a la implementaciÃ³n de la interfaz
Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function

' MÃ©todo pÃºblico que delega a la implementaciÃ³n de la interfaz
Public Function ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.Recordset
    Set ExecuteQuery = ISolicitudRepository_ExecuteQuery(sql, params)
End Function






