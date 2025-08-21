VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestSuiteResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase: CTestSuiteResult
' Proposito: Almacena los resultados de una suite de pruebas.
Option Compare Database
Option Explicit

Public SuiteName As String
Private m_Tests As Collection

Private Sub Class_Initialize()
    Set m_Tests = New Collection
End Sub

Public Sub AddTest(TestName As String, result As CTestResult)
    m_Tests.Add result, TestName
End Sub

Public Function GetResults() As Collection
    Set GetResults = m_Tests
End Function