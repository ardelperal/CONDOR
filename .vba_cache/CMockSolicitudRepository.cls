Option Compare Database
Option Explicit

' ============================================================================
' CLASE: CMockSolicitudRepository
' Arquitectura: Mock del Repositorio de Solicitudes para Pruebas Unitarias
' Version: 2.0 - Reconstrucción Total
' ============================================================================
' Mock que implementa ISolicitudRepository para aislamiento de pruebas unitarias.
' Permite verificar llamadas y configurar comportamientos de prueba.
' ============================================================================

Implements ISolicitudRepository

' ============================================================================
' PROPIEDADES PÚBLICAS PARA VERIFICACIÓN EN PRUEBAS
' ============================================================================

' Flags de llamadas
Public GetSolicitudByIdCalled As Boolean
Public SaveSolicitudCalled As Boolean

' Parámetros capturados
Public LastIdSolicitudRequested As Long
Public LastSavedSolicitud As T_Solicitud

' Valores de retorno configurables
Public GetSolicitudByIdReturnValue As T_Solicitud
Public SaveSolicitudReturnValue As Long

' Contadores
Public GetSolicitudByIdCallCount As Long
Public SaveSolicitudCallCount As Long

' ============================================================================
' CONSTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Reset
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS PARA CONFIGURACIÓN DEL MOCK
' ============================================================================

' Reinicia todas las propiedades del mock a sus valores por defecto
Public Sub Reset()
    ' Resetear flags de llamadas
    GetSolicitudByIdCalled = False
    SaveSolicitudCalled = False
    
    ' Resetear parámetros capturados
    LastIdSolicitudRequested = 0
    Set LastSavedSolicitud = Nothing
    
    ' Resetear valores de retorno
    Set GetSolicitudByIdReturnValue = Nothing
    SaveSolicitudReturnValue = 1 ' Valor por defecto exitoso
    
    ' Resetear contadores
    GetSolicitudByIdCallCount = 0
    SaveSolicitudCallCount = 0
End Sub

' Configura la solicitud que debe devolver GetSolicitudById
Public Sub SetGetSolicitudByIdReturnValue(ByVal solicitud As T_Solicitud)
    Set GetSolicitudByIdReturnValue = solicitud
End Sub

' Configura el ID que debe devolver SaveSolicitud
Public Sub SetSaveSolicitudReturnValue(ByVal returnId As Long)
    SaveSolicitudReturnValue = returnId
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE ISolicitudRepository
' ============================================================================

' Implementación mock de GetSolicitudById
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

' Implementación mock de SaveSolicitud
Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    ' Registrar la llamada
    SaveSolicitudCalled = True
    SaveSolicitudCallCount = SaveSolicitudCallCount + 1
    Set LastSavedSolicitud = solicitud
    
    ' Devolver el valor configurado
    ISolicitudRepository_SaveSolicitud = SaveSolicitudReturnValue
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (DELEGACIÓN)
' ============================================================================

' Método público que delega a la implementación de la interfaz
Public Function GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    Set GetSolicitudById = ISolicitudRepository_GetSolicitudById(idSolicitud)
End Function

' Método público que delega a la implementación de la interfaz
Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function






