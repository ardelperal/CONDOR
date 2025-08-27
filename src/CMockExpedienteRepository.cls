VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockExpedienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' CLASE: CMockExpedienteRepository
' Arquitectura: Mock del Repositorio de Expedientes para Pruebas Unitarias
' Version: 1.0 - Creación Inicial
' ============================================================================
' Mock que implementa IExpedienteRepository para aislamiento de pruebas unitarias.
' Permite verificar llamadas y configurar comportamientos de prueba.
' ============================================================================

Implements IExpedienteRepository

' ============================================================================
' PROPIEDADES PÚBLICAS PARA VERIFICACIÓN EN PRUEBAS
' ============================================================================

' Flags de llamadas
Public ObtenerExpedientePorIdCalled As Boolean
Public ObtenerExpedientePorNemotecnicoCalled As Boolean
Public ObtenerExpedientesActivosParaSelectorCalled As Boolean

' Parámetros capturados
Public LastIdExpedienteRequested As Long
Public LastNemotecnicoRequested As String

' Valores de retorno configurables
Public ObtenerExpedientePorIdReturnValue As DAO.recordset
Public ObtenerExpedientePorNemotecnicoReturnValue As DAO.recordset
Public ObtenerExpedientesActivosParaSelectorReturnValue As DAO.recordset

' Contadores
Public ObtenerExpedientePorIdCallCount As Long
Public ObtenerExpedientePorNemotecnicoCallCount As Long
Public ObtenerExpedientesActivosParaSelectorCallCount As Long

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
    ObtenerExpedientePorIdCalled = False
    ObtenerExpedientePorNemotecnicoCalled = False
    ObtenerExpedientesActivosParaSelectorCalled = False
    
    ' Resetear parámetros capturados
    LastIdExpedienteRequested = 0
    LastNemotecnicoRequested = ""
    
    ' Resetear valores de retorno
    Set ObtenerExpedientePorIdReturnValue = Nothing
    Set ObtenerExpedientePorNemotecnicoReturnValue = Nothing
    Set ObtenerExpedientesActivosParaSelectorReturnValue = Nothing
    
    ' Resetear contadores
    ObtenerExpedientePorIdCallCount = 0
    ObtenerExpedientePorNemotecnicoCallCount = 0
    ObtenerExpedientesActivosParaSelectorCallCount = 0
End Sub

' Configura el Recordset que debe devolver ObtenerExpedientePorId
Public Sub SetObtenerExpedientePorIdReturnValue(ByVal rs As DAO.recordset)
    Set ObtenerExpedientePorIdReturnValue = rs
End Sub

' Configura el Recordset que debe devolver ObtenerExpedientePorNemotecnico
Public Sub SetObtenerExpedientePorNemotecnicoReturnValue(ByVal rs As DAO.recordset)
    Set ObtenerExpedientePorNemotecnicoReturnValue = rs
End Sub

' Configura el Recordset que debe devolver ObtenerExpedientesActivosParaSelector
Public Sub SetObtenerExpedientesActivosParaSelectorReturnValue(ByVal rs As DAO.recordset)
    Set ObtenerExpedientesActivosParaSelectorReturnValue = rs
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE IExpedienteRepository
' ============================================================================

' Implementación mock de ObtenerExpedientePorId
Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal idExpediente As Long) As DAO.recordset
    ' Registrar la llamada
    ObtenerExpedientePorIdCalled = True
    ObtenerExpedientePorIdCallCount = ObtenerExpedientePorIdCallCount + 1
    LastIdExpedienteRequested = idExpediente
    
    ' Devolver el valor configurado
    Set IExpedienteRepository_ObtenerExpedientePorId = ObtenerExpedientePorIdReturnValue
End Function

' Implementación mock de ObtenerExpedientePorNemotecnico
Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As DAO.recordset
    ' Registrar la llamada
    ObtenerExpedientePorNemotecnicoCalled = True
    ObtenerExpedientePorNemotecnicoCallCount = ObtenerExpedientePorNemotecnicoCallCount + 1
    LastNemotecnicoRequested = Nemotecnico
    
    ' Devolver el valor configurado
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = ObtenerExpedientePorNemotecnicoReturnValue
End Function

' Implementación mock de ObtenerExpedientesActivosParaSelector
Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector() As DAO.recordset
    ' Registrar la llamada
    ObtenerExpedientesActivosParaSelectorCalled = True
    ObtenerExpedientesActivosParaSelectorCallCount = ObtenerExpedientesActivosParaSelectorCallCount + 1
    
    ' Devolver el valor configurado
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = ObtenerExpedientesActivosParaSelectorReturnValue
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (DELEGACIÓN)
' ============================================================================

' Método público que delega a la implementación de la interfaz
Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As DAO.recordset
    Set ObtenerExpedientePorId = IExpedienteRepository_ObtenerExpedientePorId(idExpediente)
End Function

' Método público que delega a la implementación de la interfaz
Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As DAO.recordset
    Set ObtenerExpedientePorNemotecnico = IExpedienteRepository_ObtenerExpedientePorNemotecnico(Nemotecnico)
End Function

' Método público que delega a la implementación de la interfaz
Public Function ObtenerExpedientesActivosParaSelector() As DAO.recordset
    Set ObtenerExpedientesActivosParaSelector = IExpedienteRepository_ObtenerExpedientesActivosParaSelector()
End Function

