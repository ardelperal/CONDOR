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
Implements IExpedienteRepository

Private m_mockExpediente As EExpediente
Private m_mockExpedientes As Scripting.Dictionary

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' Método para que los tests configuren el expediente que debe devolver el mock
Public Sub ConfigureObtenerExpedientePorId(ByVal expediente As EExpediente)
    Set m_mockExpediente = expediente
End Sub

' Método para que los tests configuren el expediente que debe devolver el mock por nemotécnico
Public Sub ConfigureObtenerExpedientePorNemotecnico(ByVal expediente As EExpediente)
    Set m_mockExpediente = expediente
End Sub

' Método para que los tests configuren la colección de expedientes que debe devolver el mock
Public Sub ConfigureObtenerExpedientesActivosParaSelector(ByVal expedientes As Scripting.Dictionary)
    Set m_mockExpedientes = expedientes
End Sub

' Resetea el estado del mock para asegurar el aislamiento entre tests
Public Sub Reset()
    Set m_mockExpediente = Nothing
    Set m_mockExpedientes = Nothing
End Sub

Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    ' Devuelve el expediente configurado por el test
    Set IExpedienteRepository_ObtenerExpedientePorId = m_mockExpediente
End Function

Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    ' Devuelve el expediente configurado por el test
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = m_mockExpediente
End Function

Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector() As Scripting.Dictionary
    ' Devuelve la colección de expedientes configurada por el test
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = m_mockExpedientes
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para ObtenerExpedientePorId
Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    Set ObtenerExpedientePorId = IExpedienteRepository_ObtenerExpedientePorId(idExpediente)
End Function

' Método público de conveniencia para ObtenerExpedientePorNemotecnico
Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set ObtenerExpedientePorNemotecnico = IExpedienteRepository_ObtenerExpedientePorNemotecnico(Nemotecnico)
End Function

' Método público de conveniencia para ObtenerExpedientesActivosParaSelector
Public Function ObtenerExpedientesActivosParaSelector() As Scripting.Dictionary
    Set ObtenerExpedientesActivosParaSelector = IExpedienteRepository_ObtenerExpedientesActivosParaSelector()
End Function