Attribute VB_Name = "modTestRunner"
Option Compare Database
Option Explicit

' ============================================================================
' ¡¡¡ REQUISITO DE COMPILACIÓN CRÍTICO !!!
' Este módulo utiliza el descubrimiento automático de pruebas a través del objeto
' Application.VBE. Para que el proyecto compile, es OBLIGATORIO tener
' activada la referencia a la librería:
' "Microsoft Visual Basic for Applications Extensibility 5.3"
' (Herramientas -> Referencias -> Marcar la casilla correspondiente)
' Si esta referencia falta, el proyecto NO COMPILARÁ.
' ============================================================================

' Variable de estado global para almacenar las suites registradas
Private m_SuiteNames As Object

' ============================================================================
' PUNTO ÚNICO DE APROVISIONAMIENTO PREVIO
' ============================================================================

Public Sub ResetTestEnvironment()
    ' Punto único de reseteo/provisión del entorno de testing.
    ' Garantiza ejecución desatendida sin diálogos.
    On Error GoTo ErrorHandler
    
    ' Configurar Access en modo completamente silencioso
    Application.Echo False
    DoCmd.Echo False
    DoCmd.SetWarnings False
    
    ' Preflight del workspace ANTES de ejecutar suites
    On Error Resume Next
    Call modTestUtils.EnsureFolder(modTestUtils.GetWorkspacePath())
    On Error GoTo 0
    modTestUtils.PrepareCoreTestDatabases
    On Error GoTo ErrorHandler
    
    ' Asegurar que las carpetas del entorno de testing existen
    Dim workspacePath As String
    workspacePath = modTestUtils.GetWorkspacePath()
    modTestUtils.EnsureFolder workspacePath
    
    ' Asegurar que la carpeta de fixtures existe
    Dim fixturesPath As String
    fixturesPath = modTestUtils.GetDatabaseFixturesPath()
    modTestUtils.EnsureFolder fixturesPath
    
    ' Aprovisiona SOLICITUD si falta (idempotente)
    Dim ws As String: ws = modTestUtils.GetWorkspacePath()
    Dim solicitudDb As String: solicitudDb = modTestUtils.JoinPath(ws, "Solicitud_integration_test.accdb")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(solicitudDb) Then
        Call modTestUtils.PrepareTestDatabase("Solicitud_test_template.accdb", "Solicitud_integration_test.accdb")
    End If
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "modTestRunner.ResetTestEnvironment", "Error en aprovisionamiento previo: " & Err.Description
End Sub

' Función para compatibilidad con CLI (debe estar fuera del bloque condicional)
Public Function RunAllTests() As String
    On Error GoTo ErrorHandler
    
    ' PUNTO ÚNICO DE APROVISIONAMIENTO PREVIO
    Call ResetTestEnvironment
    Call modTestUtils.CloseAllWordInstancesForTesting
    
    ' Crear el ErrorHandler para el PROPIO RUNNER usando la configuración Singleton
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Inicializar el diccionario de suites
    Set m_SuiteNames = New Scripting.Dictionary
    m_SuiteNames.CompareMode = TextCompare
    
    ' Descubrir y registrar suites
    Call DiscoverAndRegisterSuites
    
    Dim reporter As ITestReporter
    Dim reporterImpl As New CTestReporter
    Set reporter = reporterImpl
    
    Dim allResults As Object
    Set allResults = ExecuteAllSuites(m_SuiteNames, errorHandler)
    
    reporter.Initialize allResults
    
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    ' Limpieza final de Word tras completar todos los tests
    Call modTestUtils.CloseAllWordInstancesForTesting
    Call modTestUtils.KillRecentWordProcesses(10, True)
    
    RunAllTests = reportString
    
    Exit Function
    
ErrorHandler:
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modTestRunner.RunAllTests", True
    End If
    RunAllTests = "FALLO CRÍTICO EN EL MOTOR DE PRUEBAS: " & Err.Description & vbCrLf & "RESULT: FAILED"
End Function

' Función para registrar suites manualmente (sin dependencia de VBE)
' Usada por ExecuteAllTestsForCLI para evitar problemas con referencias externas
Private Sub RegisterSuitesManually()
    On Error Resume Next
    
    ' Inicializar el diccionario si no existe
    If m_SuiteNames Is Nothing Then
        Set m_SuiteNames = New Scripting.Dictionary
        m_SuiteNames.CompareMode = TextCompare
    End If
    
    ' Registrar todas las suites de prueba conocidas manualmente
    ' Esto evita la dependencia de Application.VBE que puede fallar desde CLI
    
    ' Suites de prueba unitaria
    m_SuiteNames.Add "TestModAssertRunAll", "TestModAssertRunAll"
    m_SuiteNames.Add "TestSolicitudServiceRunAll", "TestSolicitudServiceRunAll"
    m_SuiteNames.Add "TestAppManagerRunAll", "TestAppManagerRunAll"
    m_SuiteNames.Add "TestAuthServiceRunAll", "TestAuthServiceRunAll"
    m_SuiteNames.Add "TestCConfigRunAll", "TestCConfigRunAll"
    m_SuiteNames.Add "TestCExpedienteServiceRunAll", "TestCExpedienteServiceRunAll"
    m_SuiteNames.Add "TestDocumentServiceRunAll", "TestDocumentServiceRunAll"
    m_SuiteNames.Add "TestErrorHandlerServiceRunAll", "TestErrorHandlerServiceRunAll"
    m_SuiteNames.Add "TestOperationLoggerRunAll", "TestOperationLoggerRunAll"
    m_SuiteNames.Add "TestWorkflowServiceRunAll", "TestWorkflowServiceRunAll"
    
    ' Suites de prueba de integración
    m_SuiteNames.Add "TIAuthRepositoryRunAll", "TIAuthRepositoryRunAll"
    m_SuiteNames.Add "TIDocumentServiceRunAll", "TIDocumentServiceRunAll"
    m_SuiteNames.Add "TIExpedienteRepositoryRunAll", "TIExpedienteRepositoryRunAll"
    m_SuiteNames.Add "TIFileSystemRunAll", "TIFileSystemRunAll"
    m_SuiteNames.Add "TIMapeoRepositoryRunAll", "TIMapeoRepositoryRunAll"
    m_SuiteNames.Add "TINotificationServiceRunAll", "TINotificationServiceRunAll"
    m_SuiteNames.Add "TIOperationRepositoryRunAll", "TIOperationRepositoryRunAll"
    m_SuiteNames.Add "TISolicitudRepositoryRunAll", "TISolicitudRepositoryRunAll"
    m_SuiteNames.Add "TIWordManagerRunAll", "TIWordManagerRunAll"
    m_SuiteNames.Add "TIWorkflowRepositoryRunAll", "TIWorkflowRepositoryRunAll"
    
    On Error GoTo 0
End Sub

' Alias para compatibilidad con CLI
Public Function ExecuteAllTests() As String
    ExecuteAllTests = RunAllTests()
End Function

' Función específica para CLI - Sin MsgBox, manejo robusto de errores
Public Function ExecuteAllTestsForCLI() As String
    On Error GoTo ErrorHandler
    
    ' PUNTO ÚNICO DE APROVISIONAMIENTO PREVIO
    Call ResetTestEnvironment
    Call modTestUtils.CloseAllWordInstancesForTesting
    
    ' Configurar Access en modo completamente silencioso (ya hecho en ResetTestEnvironment)
    On Error Resume Next
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Crear el ErrorHandler para el PROPIO RUNNER usando la configuración Singleton
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Inicializar el diccionario de suites usando registro manual (sin VBE)
    Set m_SuiteNames = New Scripting.Dictionary
    m_SuiteNames.CompareMode = TextCompare
    
    ' Registrar suites manualmente para evitar problemas con VBE desde CLI
    Call RegisterSuitesManually
    
    ' Ejecutar todas las suites registradas
    Dim allResults As Object
    Set allResults = ExecuteAllSuites(m_SuiteNames, errorHandler)
    
    ' Generar el reporte usando el reporter estándar
    Dim reporter As ITestReporter
    Dim reporterImpl As New CTestReporter
    Set reporter = reporterImpl
    
    reporter.Initialize allResults
    
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    ' Limpieza final de Word tras completar todos los tests
    Call modTestUtils.CloseAllWordInstancesForTesting
    Call modTestUtils.KillRecentWordProcesses(10, True)
    
    ExecuteAllTestsForCLI = reportString
    GoTo Cleanup ' Salto a la limpieza en caso de éxito

ErrorHandler:
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modTestRunner.ExecuteAllTestsForCLI", True
    End If
    ExecuteAllTestsForCLI = "FALLO CRÍTICO EN EL MOTOR DE PRUEBAS CLI: " & Err.Description & vbCrLf & "RESULT: FAILURE"
    GoTo Cleanup ' Salto a la limpieza en caso de error
    
Cleanup:
    On Error Resume Next ' Blindaje final
    Debug.Print "Ejecutando limpieza final de procesos..."
    On Error GoTo 0
    ' No hay Exit Function aquí para que la función pueda devolver su valor
End Function

' MOTOR DE EJECUCIÓN DE PRUEBAS - FRAMEWORK ORIENTADO A OBJETOS
' Arquitectura: Separación de Responsabilidades (Ejecución vs. Reporte)
' Version: 3.0 - Refactorización Crítica
'******************************************************************************

'******************************************************************************
' FUNCIÓN PRINCIPAL - ORQUESTADOR DEL FRAMEWORK
'******************************************************************************

' Función principal que orquesta todo el proceso: registrar, ejecutar y reportar
Public Function RunTestFramework() As String
    On Error GoTo errorHandler
    
    ' PUNTO ÚNICO DE APROVISIONAMIENTO PREVIO
    Call ResetTestEnvironment
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Set m_SuiteNames = New Scripting.Dictionary
    m_SuiteNames.CompareMode = TextCompare
    
    ' 1. REGISTRAR
    Call DiscoverAndRegisterSuites
    
    ' 2. EJECUTAR
    Dim allResults As Object
    Set allResults = ExecuteAllSuites(m_SuiteNames, errorHandler)
    
    ' 3. GENERAR REPORTE
    Dim reporter As New CTestReporter
    reporter.Initialize allResults
    
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    ' 4. DEVOLVER (en lugar de presentar)
    RunTestFramework = reportString
    
    Exit Function
    
errorHandler:
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modTestRunner.RunTestFramework", True
    End If
    RunTestFramework = "FALLO CRÍTICO EN EL MOTOR DE PRUEBAS: " & Err.Description
End Function

'******************************************************************************
' MOTOR DE DESCUBRIMIENTO
'******************************************************************************

' Función que descubre automáticamente las suites de prueba en el proyecto VBA
Private Sub DiscoverAndRegisterSuites()
    On Error GoTo ErrorHandler
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Iterar sobre todos los componentes del proyecto VBA
    Dim comp As Object ' VBComponent
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        ' Buscar módulos que contengan funciones "RunAll"
        If comp.Type = 1 Then ' vbext_ct_StdModule
            Dim moduleName As String
            moduleName = comp.Name
            
            ' Verificar si el módulo tiene una función RunAll
            If HasRunAllFunction(moduleName) Then
                Dim suiteKey As String
                suiteKey = moduleName & "RunAll"
                
                ' Añadir al diccionario si no existe ya
                If Not m_SuiteNames.Exists(suiteKey) Then
                    m_SuiteNames.Add suiteKey, suiteKey
                End If
            End If
        End If
    Next comp
    
    Exit Sub
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.DiscoverAndRegisterSuites", True
    Err.Raise Err.Number, "modTestRunner.DiscoverAndRegisterSuites", "FALLO CRÍTICO: " & Err.Description
End Sub

' Función auxiliar para verificar si un módulo tiene una función RunAll
Private Function HasRunAllFunction(ByVal moduleName As String) As Boolean
    On Error Resume Next
    
    ' Intentar ejecutar la función RunAll del módulo
    ' Si existe, no habrá error; si no existe, habrá error
    Dim testCall As String
    testCall = moduleName & "RunAll"
    
    ' Verificar si la función existe sin ejecutarla
    ' Esto es una aproximación - en un entorno real podrías usar reflexión VBA
    HasRunAllFunction = (moduleName Like "Test*" Or moduleName Like "TI*")
    
    On Error GoTo 0
End Function

'******************************************************************************
' MOTOR DE EJECUCIÓN
'******************************************************************************

' Función que ejecuta todas las suites registradas y devuelve resultados
Private Function ExecuteAllSuites(ByVal suiteNames As Object, ByVal runnerErrorHandler As IErrorHandlerService) As Object
    Dim allResults As New Scripting.Dictionary
    allResults.CompareMode = TextCompare
    
    ' Las plantillas se verifican automáticamente en PrepareTestDatabase de cada suite
    
    Dim i As Integer
    Dim suiteKeys As Variant
    suiteKeys = suiteNames.Keys()
    
    For i = 0 To UBound(suiteKeys)
        Dim suiteName As String
        suiteName = suiteKeys(i)
        
        ' Ejecutar la suite usando Application.Run
        On Error Resume Next
        Dim suiteResult As CTestSuiteResult
        Set suiteResult = Application.Run(suiteName)
        
        If Err.Number = 0 And Not suiteResult Is Nothing Then
            allResults.Add suiteName, suiteResult
        Else
            ' Crear un resultado de error para la suite que falló
            Dim errorSuite As New CTestSuiteResult
            errorSuite.Initialize suiteName
            
            Dim errorTest As New CTestResult
            errorTest.Initialize "Suite_Execution_Error"
            errorTest.Fail "Error ejecutando suite: " & Err.Description
            
            errorSuite.AddResult errorTest
            allResults.Add suiteName, errorSuite
            
            ' Loguear el error usando el errorHandler PRE-EXISTENTE
            If Not runnerErrorHandler Is Nothing Then
                runnerErrorHandler.LogError Err.Number, Err.Description, "modTestRunner.ExecuteAllSuites", True
            End If
        End If
        
        ' Limpiar el estado de error antes de la siguiente iteración
        On Error GoTo 0
        Err.Clear
    Next i
    
    Set ExecuteAllSuites = allResults
End Function

'******************************************************************************
' FUNCIÓN DE COMPATIBILIDAD PARA EJECUCIÓN MANUAL
'******************************************************************************

' Función de compatibilidad para ejecución manual desde modAppManager
Public Sub EjecutarTodasLasPruebas()
    Call RunTestFramework
End Sub














