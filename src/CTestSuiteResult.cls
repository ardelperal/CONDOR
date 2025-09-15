Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' REQUISITO DE COMPILACIÓN:
' Este módulo utiliza Scripting.Dictionary y requiere la referencia a
' la librería "Microsoft Scripting Runtime" para compilar correctamente.

' ============================================================================
' Clase: CTestSuiteResult
' Descripción: Almacena los resultados de una suite de pruebas completa.
' Lecciones Aplicadas: 21 (Framework Robusto), 41 (Scripting.Dictionary)
' ============================================================================

Public Name As String
Public Results As Object ' Clave: Nombre del test, Valor: CTestResult
Public TotalTests As Long
Public PassedTests As Long
Public FailedTests As Long

Private Sub Class_Initialize()
    Set Results = New Scripting.Dictionary
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
End Sub

' Propiedad que encapsula la lógica de si la suite ha pasado por completo
Public Property Get AllTestsPassed() As Boolean
    If TotalTests > 0 And FailedTests = 0 Then
        AllTestsPassed = True
    Else
        AllTestsPassed = False
    End If
End Property

' Método de inicialización robusto (Lección 21)
Public Sub Initialize(ByVal suiteName As String)
    Name = suiteName
End Sub

' Método crítico que faltaba
Public Sub AddResult(ByVal testResult As CTestResult)
    If Results.Exists(testResult.Name) Then
        ' Evitar duplicados, aunque no debería ocurrir
        Results.Remove testResult.Name
    End If
    
    Results.Add testResult.Name, testResult
    
    TotalTests = TotalTests + 1
    If testResult.Passed Then
        PassedTests = PassedTests + 1
    Else
        FailedTests = FailedTests + 1
    End If
End Sub