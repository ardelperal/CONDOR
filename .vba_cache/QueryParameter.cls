VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Clase auxiliar para encapsular parámetros de consulta SQL
' Versión 1.0 - Soporte para consultas parametrizadas seguras

Private m_name As String
Private m_value As Variant

' Constructor
Public Sub Initialize(ByVal paramName As String, ByVal paramValue As Variant)
    m_name = paramName
    m_value = paramValue
End Sub

' Propiedades
Public Property Get Name() As String
    Name = m_name
End Property

Public Property Get Value() As Variant
    If IsObject(m_value) Then
        Set Value = m_value
    Else
        Value = m_value
    End If
End Property