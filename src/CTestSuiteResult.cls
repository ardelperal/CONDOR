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
' Arquitectura: Gestión Orientada a Objetos de Resultados
' Version: 3.0 - Refactorización Crítica
' Fecha: 2025-01-14
' ============================================================================

' Propiedades privadas
Private m_suiteName As String
Private m_testResults As Scripting.Dictionary

' ============================================================================
' INICIALIZACIÓN
' ============================================================================

Private Sub Class_Initialize()
    Set m_testResults = New Scripting.Dictionary
    m_testResults.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set m_testResults = Nothing
End Sub

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

' Propiedad para el nombre de la suite
Public Property Get suiteName() As String
    suiteName = m_suiteName
End Property

Public Property Let suiteName(ByVal value As String)
    m_suiteName = value
End Property

' Propiedades calculadas en tiempo real
Public Property Get TotalTests() As Long
    TotalTests = m_testResults.Count
End Property

Public Property Get TestsPassed() As Long
    Dim passedCount As Long
    Dim testResult As CTestResult
    Dim key As Variant
    
    For Each key In m_testResults.Keys()
        Set testResult = m_testResults(key)
        If testResult.Success Then
            passedCount = passedCount + 1
        End If
    Next key
    
    TestsPassed = passedCount
End Property

Public Property Get TestsFailed() As Long
    TestsFailed = TotalTests - TestsPassed
End Property

' ============================================================================
' MÉTODOS PÚBLICOS
' ============================================================================

' Método de inicialización pública
Public Sub Initialize(ByVal suiteName As String)
    m_suiteName = suiteName
    Set m_testResults = New Scripting.Dictionary
    m_testResults.CompareMode = TextCompare
End Sub

' Método para añadir un resultado de prueba
Public Sub AddTestResult(ByVal result As CTestResult)
    If result Is Nothing Then
        Err.Raise 91, "CTestSuiteResult.AddTestResult", "El resultado de prueba no puede ser Nothing"
    End If
    
    m_testResults.Add result.TestName, result
End Sub

' Método alias para compatibilidad con modTestRunner
Public Sub AddTest(ByVal result As CTestResult)
    Call AddTestResult(result)
End Sub

' Función que devuelve el reporte formateado para esta suite
Public Function GetFormattedResult() As String
    Dim report As String
    
    ' Encabezado de la suite
    report = "=== SUITE: " & m_suiteName & " ===" & vbCrLf
    
    ' Detalles de cada prueba
    Dim testResult As CTestResult
    For Each testResult In m_testResults
        If testResult.Success Then
            report = report & "[? PASS] " & testResult.testName & vbCrLf
        Else
            report = report & "[? FAIL] " & testResult.testName
            If testResult.errorMessage <> "" Then
                report = report & " - " & testResult.errorMessage
            End If
            report = report & vbCrLf
        End If
    Next testResult
    
    ' Resumen de la suite
    report = report & "Resumen " & m_suiteName & ": " & TestsPassed & "/" & TotalTests & " pruebas exitosas" & vbCrLf
    
    GetFormattedResult = report
End Function

' ============================================================================
' MÉTODOS DE ACCESO A RESULTADOS INDIVIDUALES
' ============================================================================

' Método para obtener un resultado específico por clave
Public Function GetTestResult(ByVal testName As String) As CTestResult
    If Not m_testResults.Exists(testName) Then
        Err.Raise 9, "CTestSuiteResult.GetTestResult", "Test no encontrado: " & testName
    End If
    
    Set GetTestResult = m_testResults(testName)
End Function

' Método para verificar si la suite completa pasó
Public Function AllTestsPassed() As Boolean
    AllTestsPassed = (TestsFailed = 0 And TotalTests > 0)
End Function

' Método para obtener lista de pruebas fallidas
Public Function GetFailedTestNames() As String
    Dim failedNames As String
    Dim testResult As CTestResult
    
    For Each testResult In m_testResults
        If Not testResult.Success Then
            If failedNames <> "" Then failedNames = failedNames & ", "
            failedNames = failedNames & testResult.testName
        End If
    Next testResult
    
    GetFailedTestNames = failedNames
End Function


