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


' =====================================================
' INTERFAZ: IConfig
' DESCRIPCIÓN: Contrato para el servicio de configuración.
' Versión: 3.0 - Sincronizada con la implementación CConfig
' =====================================================

Public Function GetValue(ByVal key As String) As Variant
End Function

Public Function HasKey(ByVal key As String) As Boolean
End Function

Public Function GetDataPath() As String
End Function

Public Function GetDatabasePassword() As String
End Function

Public Function GetAttachmentsPath() As String
End Function

Public Function GetCorreosDBPath() As String
End Function

Public Function GetUsuarioActual() As String
End Function

Public Function GetCorreoAdministrador() As String
End Function

Public Function GetIDAplicacionCondor() As String
End Function

Public Function GetLanzaderaDataPath() As String
End Function

Public Function GetLanzaderaPassword() As String
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
End Sub


