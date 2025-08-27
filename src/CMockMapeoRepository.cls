VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockMapeoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' CLASE: CMockMapeoRepository
' Arquitectura: Mock del Repositorio de Mapeo para Pruebas Unitarias
' Version: 1.0 - Creación Inicial
' ============================================================================
' Mock que implementa IMapeoRepository para aislamiento de pruebas unitarias.
' Permite verificar llamadas y configurar comportamientos de prueba.
' ============================================================================

Implements IMapeoRepository

' ============================================================================
' PROPIEDADES PÚBLICAS PARA VERIFICACIÓN EN PRUEBAS
' ============================================================================

' Flags de llamadas
Public GetMapeoPorTipoCalled As Boolean

' Parámetros capturados
Public LastTipoSolicitudRequested As String

' Valores de retorno configurables
Public GetMapeoPorTipoReturnValue As DAO.recordset

' Contadores
Public GetMapeoPorTipoCallCount As Long

' ============================================================================
' MÉTODOS PÚBLICOS PARA CONFIGURACIÓN DEL MOCK
' ============================================================================

' Reinicia el estado del mock para una nueva prueba
Public Sub Reset()
    GetMapeoPorTipoCalled = False
    LastTipoSolicitudRequested = ""
    Set GetMapeoPorTipoReturnValue = Nothing
    GetMapeoPorTipoCallCount = 0
End Sub

' Configura el Recordset que debe devolver GetMapeoPorTipo
Public Sub SetGetMapeoPorTipoReturnValue(ByVal rs As DAO.recordset)
    Set GetMapeoPorTipoReturnValue = rs
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE IMapeoRepository
' ============================================================================

' Implementación mock de GetMapeoPorTipo
Private Function IMapeoRepository_GetMapeoPorTipo(ByVal tipoSolicitud As String) As DAO.recordset
    ' Registrar la llamada
    GetMapeoPorTipoCalled = True
    GetMapeoPorTipoCallCount = GetMapeoPorTipoCallCount + 1
    LastTipoSolicitudRequested = tipoSolicitud
    
    ' Devolver el valor configurado
    Set IMapeoRepository_GetMapeoPorTipo = GetMapeoPorTipoReturnValue
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (DELEGACIÓN)
' ============================================================================

' Método público que delega a la implementación de la interfaz
Public Function GetMapeoPorTipo(ByVal tipoSolicitud As String) As DAO.recordset
    Set GetMapeoPorTipo = IMapeoRepository_GetMapeoPorTipo(tipoSolicitud)
End Function

