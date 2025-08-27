VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Clase: CTestReporter
' Propósito: Agrega resultados de múltiples suites y genera un informe consolidado.
' Implementa: ITestReporter

Implements ITestReporter

Private m_suiteResults As Collection

' El constructor recibe la colección de resultados del runner
Public Sub Initialize(ByVal results As Collection)
    Set m_suiteResults = results
End Sub

' --- Propiedades Calculadas Globales ---
Public Property Get TotalSuites() As Long
    TotalSuites = m_suiteResults.count
End Property

Public Property Get TotalTests() As Long
    Dim total As Long
    Dim result As CTestSuiteResult
    For Each result In m_suiteResults
        total = total + result.TotalTests
    Next result
    TotalTests = total
End Property

Public Property Get TotalPassed() As Long
    Dim total As Long
    Dim result As CTestSuiteResult
    For Each result In m_suiteResults
        total = total + result.TestsPassed
    Next result
    TotalPassed = total
End Property

Public Property Get TotalFailed() As Long
    TotalFailed = TotalTests - TotalPassed
End Property

' --- Método Principal de Generación de Informe ---
Public Function GenerateReport() As String
    Dim report As String
    Dim suiteResult As CTestSuiteResult

    report = String(80, "=") & vbCrLf
    report = report & "INFORME CONSOLIDADO DE PRUEBAS UNITARIAS" & vbCrLf
    report = report & "Fecha: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    report = report & String(80, "=") & vbCrLf & vbCrLf

    For Each suiteResult In m_suiteResults
        report = report & suiteResult.GetFormattedResult() & vbCrLf
    Next suiteResult

    report = report & String(80, "-") & vbCrLf
    report = report & "RESUMEN GLOBAL:" & vbCrLf
    report = report & "Suites ejecutadas: " & TotalSuites & vbCrLf
    report = report & "Total de pruebas: " & TotalTests & vbCrLf
    report = report & "Pruebas exitosas: " & TotalPassed & vbCrLf
    report = report & "Pruebas fallidas: " & TotalFailed & vbCrLf

    If TotalFailed = 0 Then
        report = report & "RESULTADO: ? TODAS LAS PRUEBAS PASARON" & vbCrLf
        report = report & "RESULT: SUCCESS" & vbCrLf
    Else
        report = report & "RESULTADO: ? HAY PRUEBAS FALLIDAS" & vbCrLf
        report = report & "RESULT: FAILED" & vbCrLf
    End If

    report = report & String(80, "=") & vbCrLf
    GenerateReport = report
End Function

'******************************************************************************
' IMPLEMENTACIÓN DE LA INTERFAZ ITestReporter
'******************************************************************************

' Implementación de ITestReporter.Initialize
Private Sub ITestReporter_Initialize(ByVal results As Collection)
    Initialize results
End Sub

' Implementación de ITestReporter.GenerateReport
Private Function ITestReporter_GenerateReport() As String
    ITestReporter_GenerateReport = GenerateReport()
End Function


