VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Clase: CTestResult
' Proposito: Almacena el resultado de un unico test.

Public testName As String
Public Success As Boolean
Public errorMessage As String

' Método de inicialización
Public Sub Initialize(ByVal testName As String)
    Me.testName = testName
    Me.Success = False
    Me.errorMessage = ""
End Sub

' Método para marcar la prueba como fallida
Public Sub Fail(ByVal errorMessage As String)
    Me.Success = False
    Me.errorMessage = errorMessage
End Sub

' Método para marcar la prueba como exitosa
Public Sub Pass()
    Me.Success = True
    Me.errorMessage = ""
End Sub


