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

Implements ITestReporter

Private m_suiteResults As Scripting.Dictionary

' --- MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24) ---
Public Sub Initialize(ByVal suiteResults As Scripting.Dictionary)
    Set m_suiteResults = suiteResults
End Sub

' --- IMPLEMENTACIÓN DEL CONTRATO ITestReporter (Lección 1) ---
Private Sub ITestReporter_Initialize(ByVal suiteResults As Scripting.Dictionary)
    Call Me.Initialize(suiteResults)
End Sub

Private Function ITestReporter_GenerateReport() As String
    ITestReporter_GenerateReport = Me.GenerateReport()
End Function

' --- PROPIEDADES Y MÉTODOS PÚBLICOS ---
Public Property Get TotalSuites() As Long
    TotalSuites = m_suiteResults.Count
End Property

Public Property Get TotalTests() As Long
    Dim total As Long
    Dim result As CTestSuiteResult
    Dim key As Variant
    
    For Each key In m_suiteResults.Keys()
        Set result = m_suiteResults(key)
        total = total + result.TotalTests
    Next key
    
    TotalTests = total
End Property

Public Property Get TotalPassed() As Long
    Dim total As Long
    Dim result As CTestSuiteResult
    Dim key As Variant
    
    For Each key In m_suiteResults.Keys()
        Set result = m_suiteResults(key)
        total = total + result.PassedTests
    Next key
    
    TotalPassed = total
End Property

Public Property Get TotalFailed() As Long
    TotalFailed = TotalTests - TotalPassed
End Property

Public Function GenerateReport() As String
    Dim report As String
    Dim suiteResult As CTestSuiteResult
    Dim testResult As CTestResult
    Dim key As Variant
    
    report = String(80, "=") & vbCrLf
    report = report & "INFORME CONSOLIDADO DE PRUEBAS UNITARIAS" & vbCrLf
    report = report & "Fecha: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    report = report & String(80, "=") & vbCrLf & vbCrLf
    
    For Each key In m_suiteResults.Keys()
        Set suiteResult = m_suiteResults(key)
        report = report & "SUITE: " & suiteResult.Name & " (" & suiteResult.PassedTests & "/" & suiteResult.TotalTests & " pasadas)" & vbCrLf
        report = report & String(Len("SUITE: " & suiteResult.Name), "-") & vbCrLf
        
        Dim testKey As Variant
        For Each testKey In suiteResult.Results.Keys()
            Set testResult = suiteResult.Results(testKey)
            If testResult.Passed Then
                report = report & " [V] PASS: " & testResult.Name & vbCrLf
            Else
                report = report & " [X] FAIL: " & testResult.Name & vbCrLf
                report = report & "     L--> Error: " & testResult.ErrorMessage & vbCrLf
            End If
        Next testKey
        report = report & vbCrLf
    Next key
    
    report = report & String(80, "-") & vbCrLf
    report = report & "RESUMEN GLOBAL" & vbCrLf
    report = report & "Suites ejecutadas: " & TotalSuites & vbCrLf
    report = report & "Total de pruebas: " & TotalTests & vbCrLf
    report = report & "Pruebas exitosas: " & TotalPassed & vbCrLf
    report = report & "Pruebas fallidas: " & TotalFailed & vbCrLf
    report = report & String(80, "-") & vbCrLf
    
    If TotalFailed = 0 Then
        report = report & "RESULTADO: V TODAS LAS PRUEBAS PASARON" & vbCrLf
        report = report & "RESULT: SUCCESS" & vbCrLf
    Else
        report = report & "RESULTADO: X HAY PRUEBAS FALLIDAS" & vbCrLf
        report = report & "RESULT: FAILED" & vbCrLf
    End If
    
    report = report & String(80, "=") & vbCrLf
    
    GenerateReport = report
End Function