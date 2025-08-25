VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Contrato para el servicio de configuraciÃ³n. VersiÃ³n 2.0.

Public Function GetValue(ByVal key As String) As Variant
End Function

Public Function HasKey(ByVal key As String) As Boolean
End Function

Public Function GetDataPath() As String
End Function

Public Function GetDatabasePassword() As String
End Function











