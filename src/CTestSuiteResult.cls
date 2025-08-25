VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestSuiteResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' CLASE INTELIGENTE DE RESULTADOS DE SUITE DE PRUEBAS
' Arquitectura: GestiÃ³n Orientada a Objetos de Resultados
' Version: 3.0 - RefactorizaciÃ³n CrÃ­tica
' Fecha: 2025-01-14
' ============================================================================

' Propiedades privadas
Private m_suiteName As String
Private m_testResults As Collection

' ============================================================================
' INICIALIZACIÃ“N
' ============================================================================

Private Sub Class_Initialize()
    Set m_testResults = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_testResults = Nothing
End Sub

' ============================================================================
' PROPIEDADES PÃšBLICAS
' ============================================================================

' Propiedad para el nombre de la suite
Public Property Get SuiteName() As String
    SuiteName = m_suiteName
End Property

Public Property Let SuiteName(ByVal value As String)
    m_suiteName = value
End Property

' Propiedades calculadas en tiempo real
Public Property Get TotalTests() As Long
    TotalTests = m_testResults.Count
End Property

Public Property Get TestsPassed() As Long
    Dim count As Long
    Dim testResult As CTestResult
    
    For Each testResult In m_testResults
        If testResult.Success Then
            count = count + 1
        End If
    Next testResult
    
    TestsPassed = count
End Property

Public Property Get TestsFailed() As Long
    TestsFailed = TotalTests - TestsPassed
End Property

' ============================================================================
' MÃ‰TODOS PÃšBLICOS
' ============================================================================

' MÃ©todo para aÃ±adir un resultado de prueba
Public Sub AddTestResult(ByVal result As CTestResult)
    If result Is Nothing Then
        Err.Raise 91, "CTestSuiteResult.AddTestResult", "El resultado de prueba no puede ser Nothing"
    End If
    
    m_testResults.Add result
End Sub

' FunciÃ³n que devuelve el reporte formateado para esta suite
Public Function GetFormattedResult() As String
    Dim report As String
    
    ' Encabezado de la suite
    report = "=== SUITE: " & m_suiteName & " ===" & vbCrLf
    
    ' Detalles de cada prueba
    Dim testResult As CTestResult
    For Each testResult In m_testResults
        If testResult.Success Then
            report = report & "[âœ… PASS] " & testResult.TestName & vbCrLf
        Else
            report = report & "[âŒ FAIL] " & testResult.TestName
            If testResult.ErrorMessage <> "" Then
                report = report & " - " & testResult.ErrorMessage
            End If
            report = report & vbCrLf
        End If
    Next testResult
    
    ' Resumen de la suite
    report = report & "Resumen " & m_suiteName & ": " & TestsPassed & "/" & TotalTests & " pruebas exitosas" & vbCrLf
    
    GetFormattedResult = report
End Function

' ============================================================================
' MÃ‰TODOS DE ACCESO A RESULTADOS INDIVIDUALES
' ============================================================================

' MÃ©todo para obtener un resultado especÃ­fico por Ã­ndice
Public Function GetTestResult(ByVal index As Long) As CTestResult
    If index < 1 Or index > m_testResults.Count Then
        Err.Raise 9, "CTestSuiteResult.GetTestResult", "Ãndice fuera de rango"
    End If
    
    Set GetTestResult = m_testResults(index)
End Function

' MÃ©todo para verificar si la suite completa pasÃ³
Public Function AllTestsPassed() As Boolean
    AllTestsPassed = (TestsFailed = 0 And TotalTests > 0)
End Function

' MÃ©todo para obtener lista de pruebas fallidas
Public Function GetFailedTestNames() As String
    Dim failedNames As String
    Dim testResult As CTestResult
    
    For Each testResult In m_testResults
        If Not testResult.Success Then
            If failedNames <> "" Then failedNames = failedNames & ", "
            failedNames = failedNames & testResult.TestName
        End If
    Next testResult
    
    GetFailedTestNames = failedNames
End Function
