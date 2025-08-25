VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase: CTestResult
' Proposito: Almacena el resultado de un unico test.
Option Compare Database
Option Explicit

Public TestName As String
Public Success As Boolean
Public ErrorMessage As String

' Método de inicialización
Public Sub Initialize(ByVal testName As String)
    Me.TestName = testName
    Me.Success = False
    Me.ErrorMessage = ""
End Sub

' Método para marcar la prueba como fallida
Public Sub Fail(ByVal errorMessage As String)
    Me.Success = False
    Me.ErrorMessage = errorMessage
End Sub

' Método para marcar la prueba como exitosa
Public Sub Pass()
    Me.Success = True
    Me.ErrorMessage = ""
End Sub
