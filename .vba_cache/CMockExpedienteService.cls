VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' CLASE: CMockExpedienteService
' Arquitectura: Mock del Servicio de Expedientes para Pruebas Unitarias
' Version: 1.0 - Creación Inicial
' ============================================================================
' Mock que implementa IExpedienteService para aislamiento de pruebas unitarias.
' Permite verificar llamadas y configurar comportamientos de prueba.
' ============================================================================

Implements IExpedienteService

' ============================================================================
' PROPIEDADES PÚBLICAS PARA VERIFICACIÓN EN PRUEBAS
' ============================================================================

' Flags de llamadas
Public GetExpedienteByIdCalled As Boolean
Public GetExpedienteByNemotecnicoCalled As Boolean
Public GetExpedientesParaSelectorCalled As Boolean

' Parámetros capturados
Public LastIdExpedienteRequested As Long
Public LastNemotecnicoRequested As String

' Valores de retorno configurables
Public GetExpedienteByIdReturnValue As EExpediente
Public GetExpedienteByNemotecnicoReturnValue As EExpediente
Public GetExpedientesParaSelectorReturnValue As DAO.recordset

' Contadores
Public GetExpedienteByIdCallCount As Long
Public GetExpedienteByNemotecnicoCallCount As Long
Public GetExpedientesParaSelectorCallCount As Long

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
    GetExpedienteByIdCalled = False
    GetExpedienteByNemotecnicoCalled = False
    GetExpedientesParaSelectorCalled = False
    
    ' Resetear parámetros capturados
    LastIdExpedienteRequested = 0
    LastNemotecnicoRequested = ""
    
    ' Resetear valores de retorno
    Set GetExpedienteByIdReturnValue = Nothing
    Set GetExpedienteByNemotecnicoReturnValue = Nothing
    Set GetExpedientesParaSelectorReturnValue = Nothing
    
    ' Resetear contadores
    GetExpedienteByIdCallCount = 0
    GetExpedienteByNemotecnicoCallCount = 0
    GetExpedientesParaSelectorCallCount = 0
End Sub

' Configura el Expediente que debe devolver GetExpedienteById
Public Sub ConfigureGetExpedienteById(ByVal expediente As EExpediente)
    Set GetExpedienteByIdReturnValue = expediente
End Sub

' Configura el Expediente que debe devolver GetExpedienteByNemotecnico
Public Sub ConfigureGetExpedienteByNemotecnico(ByVal expediente As EExpediente)
    Set GetExpedienteByNemotecnicoReturnValue = expediente
End Sub

' Configura el Recordset que debe devolver GetExpedientesParaSelector
Public Sub ConfigureGetExpedientesParaSelector(ByVal rs As DAO.recordset)
    Set GetExpedientesParaSelectorReturnValue = rs
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE IExpedienteService
' ============================================================================

' Implementación mock de GetExpedienteById
Private Function IExpedienteService_GetExpedienteById(ByVal idExpediente As Long) As EExpediente
    ' Registrar la llamada
    GetExpedienteByIdCalled = True
    GetExpedienteByIdCallCount = GetExpedienteByIdCallCount + 1
    LastIdExpedienteRequested = idExpediente
    
    ' Devolver el valor configurado
    Set IExpedienteService_GetExpedienteById = GetExpedienteByIdReturnValue
End Function

' Implementación mock de GetExpedienteByNemotecnico
Private Function IExpedienteService_GetExpedienteByNemotecnico(ByVal Nemotecnico As String) As EExpediente
    ' Registrar la llamada
    GetExpedienteByNemotecnicoCalled = True
    GetExpedienteByNemotecnicoCallCount = GetExpedienteByNemotecnicoCallCount + 1
    LastNemotecnicoRequested = Nemotecnico
    
    ' Devolver el valor configurado
    Set IExpedienteService_GetExpedienteByNemotecnico = GetExpedienteByNemotecnicoReturnValue
End Function

' Implementación mock de GetExpedientesParaSelector
Private Function IExpedienteService_GetExpedientesParaSelector() As DAO.recordset
    ' Registrar la llamada
    GetExpedientesParaSelectorCalled = True
    GetExpedientesParaSelectorCallCount = GetExpedientesParaSelectorCallCount + 1
    
    ' Devolver el valor configurado
    Set IExpedienteService_GetExpedientesParaSelector = GetExpedientesParaSelectorReturnValue
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (DELEGACIÓN)
' ============================================================================

' Método público que delega a la implementación de la interfaz
Public Function GetExpedienteById(ByVal idExpediente As Long) As EExpediente
    Set GetExpedienteById = IExpedienteService_GetExpedienteById(idExpediente)
End Function

' Método público que delega a la implementación de la interfaz
Public Function GetExpedienteByNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set GetExpedienteByNemotecnico = IExpedienteService_GetExpedienteByNemotecnico(Nemotecnico)
End Function

' Método público que delega a la implementación de la interfaz
Public Function GetExpedientesParaSelector() As DAO.recordset
    Set GetExpedientesParaSelector = IExpedienteService_GetExpedientesParaSelector()
End Function