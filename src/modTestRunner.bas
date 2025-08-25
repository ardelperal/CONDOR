Option Compare Database
Option Explicit

' ColecciÃ³n privada para registrar nombres de funciones de suite
Private m_registeredSuites As Collection

' FunciÃ³n para compatibilidad con CLI (debe estar fuera del bloque condicional)
Public Sub ExecuteAllTests()
#If DEV_MODE Then
    Call EjecutarTodasLasPruebas
#End If
End Sub

Public Sub RunTests()
#If DEV_MODE Then
    Call EjecutarTodasLasPruebas
#End If
End Sub

#If DEV_MODE Then

' ============================================================================
' MOTOR DE EJECUCIÃ“N DE PRUEBAS - FRAMEWORK ORIENTADO A OBJETOS
' Arquitectura: SeparaciÃ³n de Responsabilidades (EjecuciÃ³n vs. Reporte)
' Version: 3.0 - RefactorizaciÃ³n CrÃ­tica
' Fecha: 2025-01-14
' ============================================================================

' ============================================================================
' FUNCIÃ“N PRINCIPAL - ORQUESTADOR DEL FRAMEWORK
' ============================================================================

' FunciÃ³n principal que orquesta todo el proceso: registrar, ejecutar y reportar
Public Sub EjecutarTodasLasPruebas()
    RegisterTestSuites

    Dim suiteResults As Collection
    Set suiteResults = RunRegisteredSuites()

    ' Crear, usar y destruir el reporter
    Dim reporter As CTestReporter
    Set reporter = New CTestReporter
    reporter.Initialize suiteResults

    Debug.Print reporter.GenerateReport

    Set reporter = Nothing
End Sub

' ============================================================================
' GESTIÃ“N DE REGISTRO DE SUITES
' ============================================================================

' FunciÃ³n que registra todas las funciones de suite disponibles
Private Sub RegisterTestSuites()
    Set m_registeredSuites = New Collection
    
    ' ========== PRUEBAS UNITARIAS ==========
    ' Servicios y componentes con mocks
    m_registeredSuites.Add "Test_CConfig_RunAll"
    m_registeredSuites.Add "Test_CAuthService_RunAll"
    m_registeredSuites.Add "Test_CExpedienteService_RunAll"
    m_registeredSuites.Add "Test_OperationLogger_RunAll"
    m_registeredSuites.Add "Test_Solicitud_RunAll"
    m_registeredSuites.Add "Test_WorkflowRepository_RunAll"
    m_registeredSuites.Add "Test_AppManager_RunAll"
    m_registeredSuites.Add "Test_CSolicitudRepository_RunAll"
    m_registeredSuites.Add "Test_DocumentServiceFactory_RunAll"
    m_registeredSuites.Add "Test_ExpedienteServiceFactory_RunAll"
    m_registeredSuites.Add "Test_LoggingService_RunAll"
    
    ' ========== PRUEBAS DE INTEGRACIÓN ==========
    ' Repositorios que acceden a la base de datos real
    m_registeredSuites.Add "IntegrationTest_CExpedienteRepository_RunAll"
End Sub

' ============================================================================
' MOTOR DE EJECUCIÃ“N
' ============================================================================

' FunciÃ³n que ejecuta todas las suites registradas y devuelve resultados
Private Function RunRegisteredSuites() As Collection
    Dim results As New Collection
    Dim suiteName As Variant

    For Each suiteName In m_registeredSuites
        Dim suiteResult As CTestSuiteResult
        On Error Resume Next ' Captura si una suite individual falla
        Set suiteResult = Application.Run(CStr(suiteName))
        If Err.Number <> 0 Then
            Set suiteResult = CreateErrorSuiteResult(CStr(suiteName), Err.Description)
            Err.Clear
        End If
        On Error GoTo 0
        results.Add suiteResult
    Next suiteName

    Set RunRegisteredSuites = results
End Function

Private Function CreateErrorSuiteResult(ByVal suiteName As String, ByVal errorDesc As String) As CTestSuiteResult
    Dim errorResult As New CTestSuiteResult
    errorResult.SuiteName = suiteName

    Dim testResult As New CTestResult
    testResult.TestName = "SuiteExecutionError"
    testResult.Success = False
    testResult.ErrorMessage = "Error fatal al ejecutar la suite: " & errorDesc
    errorResult.AddTestResult testResult

    Set CreateErrorSuiteResult = errorResult
End Function

#End If









