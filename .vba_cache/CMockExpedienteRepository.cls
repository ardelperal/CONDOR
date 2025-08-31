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

Private m_mockRecordset As DAO.Recordset

' Método para que los tests configuren el recordset que debe devolver el mock
Public Sub SetObtenerExpedientePorIdReturnValue(ByVal rs As DAO.Recordset)
    Set m_mockRecordset = rs
End Sub

' Resetea el estado del mock para asegurar el aislamiento entre tests
Public Sub Reset()
    Set m_mockRecordset = Nothing
End Sub

Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal idExpediente As Long) As DAO.Recordset
    ' Devuelve el recordset configurado por el test
    Set IExpedienteRepository_ObtenerExpedientePorId = m_mockRecordset
End Function

Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As DAO.Recordset
    ' Devuelve el recordset configurado (se puede mejorar en el futuro si es necesario)
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = m_mockRecordset
End Function

Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector() As DAO.Recordset
    ' Devuelve el recordset configurado (se puede mejorar en el futuro si es necesario)
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = m_mockRecordset
End Function