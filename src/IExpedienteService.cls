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


Public Function GetExpedienteById(ByVal idExpediente As Long) As EExpediente
End Function

Public Function GetExpedientesParaSelector() As Scripting.Dictionary
End Function

Public Function GetExpedienteByNemotecnico(ByVal Nemotecnico As String) As EExpediente
End Function













