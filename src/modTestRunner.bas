Attribute VB_Name = "modTestRunner"

' Colección privada para registrar nombres de funciones de suite
Private m_SuiteNames As Collection

' Función para compatibilidad con CLI (debe estar fuera del bloque condicional)
Public Function RunAllTests() As String
    ' Ejecutar todas las pruebas y devolver resultado como string
    Dim reporter As New CTestReporter
    Dim allResults As Collection
    
    Set allResults = ExecuteAllSuites()
    
    ' Generar reporte en formato string
    Dim reportString As String
    reportString = reporter.GenerateReport(allResults)
    
    RunAllTests = reportString
End Function

'******************************************************************************
' MOTOR DE EJECUCIÓN DE PRUEBAS - FRAMEWORK ORIENTADO A OBJETOS
' Arquitectura: Separación de Responsabilidades (Ejecución vs. Reporte)
' Version: 3.0 - Refactorización Crítica
'******************************************************************************

'******************************************************************************
' FUNCIÓN PRINCIPAL - ORQUESTADOR DEL FRAMEWORK
'******************************************************************************

' Función principal que orquesta todo el proceso: registrar, ejecutar y reportar
Public Sub RunTestFramework()
    ' Inicializar colección de suites
    Set m_SuiteNames = New Collection
    
    ' Registrar todas las suites disponibles
    RegisterAllSuites
    
    ' Ejecutar todas las suites y obtener resultados
    Dim allResults As Collection
    Set allResults = ExecuteAllSuites()
    
    ' Generar y mostrar reporte
    Dim reporter As New CTestReporter
    reporter.ShowReport allResults
End Sub

'******************************************************************************
' GESTIÓN DE REGISTRO DE SUITES
'******************************************************************************

' Función que registra todas las funciones de suite disponibles
Private Sub RegisterAllSuites()
    ' Registrar todas las suites de prueba disponibles
    ' Cada suite debe seguir el patrón: Test_[ModuleName]_RunAll
    
    m_SuiteNames.Add "Test_CConfig_RunAll"
    m_SuiteNames.Add "Test_AuthService_RunAll"
    m_SuiteNames.Add "Test_CExpedienteService_RunAll"
    m_SuiteNames.Add "Test_Solicitud_RunAll"
    m_SuiteNames.Add "Test_CSolicitudRepository_RunAll"
    m_SuiteNames.Add "Test_DocumentService_RunAll"
    m_SuiteNames.Add "Test_DocumentServiceFactory_RunAll"
    m_SuiteNames.Add "Test_ExpedienteServiceFactory_RunAll"
    m_SuiteNames.Add "Test_LoggingService_RunAll"
    m_SuiteNames.Add "Test_NotificationService_RunAll"
    m_SuiteNames.Add "Test_OperationLogger_RunAll"
    m_SuiteNames.Add "Test_WordManager_RunAll"
    m_SuiteNames.Add "Test_WorkflowRepository_RunAll"
    m_SuiteNames.Add "Test_ErrorHandlerService_RunAll"
    m_SuiteNames.Add "Test_AppManager_RunAll"
End Sub

'******************************************************************************
' MOTOR DE EJECUCIÓN
'******************************************************************************

' Función que ejecuta todas las suites registradas y devuelve resultados
Private Function ExecuteAllSuites() As Collection
    Dim allResults As New Collection
    Dim i As Integer
    
    For i = 1 To m_SuiteNames.Count
        Dim suiteName As String
        suiteName = m_SuiteNames(i)
        
        ' Ejecutar la suite usando Application.Run
        On Error Resume Next
        Dim suiteResult As CTestSuiteResult
        Set suiteResult = Application.Run(suiteName)
        
        If Err.Number = 0 And Not suiteResult Is Nothing Then
            allResults.Add suiteResult
        Else
            ' Crear un resultado de error para suites que fallan
            Dim errorSuite As New CTestSuiteResult
            errorSuite.Initialize suiteName
            
            Dim errorTest As New CTestResult
            errorTest.Initialize "Suite_Execution_Error"
            errorTest.Fail "Error ejecutando suite: " & Err.Description
            
            errorSuite.AddTest errorTest
            allResults.Add errorSuite
        End If
        
        On Error GoTo 0
    Next i
    
    Set ExecuteAllSuites = allResults
End Function









