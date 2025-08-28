VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Function GetExpedienteById(ByVal idExpediente As Long) As E_Expediente
End Function

Public Function GetExpedientesParaSelector() As DAO.Recordset
End Function

Public Function GetExpedienteByNemotecnico(ByVal Nemotecnico As String) As E_Expediente
End Function













