VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOperationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Contrato para la persistencia de logs de operaciones.

Public Sub SaveLog(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
End Sub




