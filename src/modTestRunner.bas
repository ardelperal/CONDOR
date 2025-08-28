Attribute VB_Name = "modTestRunner"
Option Compare Database
Option Explicit


' Colección privada para registrar nombres de funciones de suite
Private m_SuiteNames As Collection

' Función para compatibilidad con CLI (debe estar fuera del bloque condicional)
Public Function RunAllTests() As String
    On Error GoTo ErrorHandler
    
    ' Obtener instancia del manejador de errores
    Dim errorHandler As IErrorHandlerService
    Dim config As New CConfig
    Dim fileSystem As New CFileSystem
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem) ' Created here for top-level entry point
    
    ' Inicializar colección de suites
    Set m_SuiteNames = New Collection
    
    ' Descubrir y registrar automáticamente todas las suites disponibles
    DiscoverAndRegisterSuites
    
    ' Ejecutar todas las pruebas y devolver resultado como string
    Dim reporter As ITestReporter
    Dim reporterImpl As New CTestReporter
    Set reporter = reporterImpl
    
    Dim allResults As Collection
    Set allResults = ExecuteAllSuites
    
    ' Inicializar el reportero con los resultados
    reporter.Initialize allResults
    
    ' Generar reporte en formato string
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    RunAllTests = reportString
    Exit Function
    
ErrorHandler:
    ' Usar el manejador de errores creado al inicio
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.RunAllTests", True ' Mark as critical
    
    RunAllTests = "FALLO CRÍTICO EN EL MOTOR DE PRUEBAS: " & Err.Description & vbCrLf & "RESULT: FAILED"
End Function


' Alias para compatibilidad con CLI
Public Function ExecuteAllTests() As String
    ExecuteAllTests = RunAllTests()
End Function

' Función específica para CLI - Sin MsgBox, manejo robusto de errores
Public Function ExecuteAllTestsForCLI() As String
    On Error GoTo ErrorHandler
    
    ' Obtener instancia del manejador de errores
    Dim errorHandler As IErrorHandlerService
    Dim config As New CConfig
    Dim fileSystem As New CFileSystem
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    
    ' Inicializar colección de suites
    Set m_SuiteNames = New Collection
    
    ' Descubrir y registrar automáticamente todas las suites disponibles
    DiscoverAndRegisterSuites
    
    ' Ejecutar todas las pruebas y devolver resultado como string
    Dim reporter As ITestReporter
    Dim reporterImpl As New CTestReporter
    Set reporter = reporterImpl
    
    Dim allResults As Collection
    Set allResults = ExecuteAllSuites
    
    ' Inicializar el reportero con los resultados
    reporter.Initialize allResults
    
    ' Generar reporte en formato string
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    ' Verificar si todas las pruebas pasaron
    Dim allPassed As Boolean
    allPassed = True
    
    Dim suiteResult As CTestSuiteResult
    For Each suiteResult In allResults
        If Not suiteResult.AllTestsPassed Then
            allPassed = False
            Exit For
        End If
    Next
    
    ' Añadir línea de resultado final
    If allPassed Then
        reportString = reportString & vbCrLf & "RESULT: SUCCESS"
    Else
        reportString = reportString & vbCrLf & "RESULT: FAILURE"
    End If
    
    ExecuteAllTestsForCLI = reportString
    Exit Function
    
ErrorHandler:
    ' Usar el manejador de errores para logging
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modTestRunner.ExecuteAllTestsForCLI", True
    End If
    
    ExecuteAllTestsForCLI = "FALLO CRÍTICO EN EL MOTOR DE PRUEBAS CLI: " & Err.Description & vbCrLf & "RESULT: FAILURE"
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
    On Error GoTo ErrorHandler
    
    ' Obtener instancia del manejador de errores
    Dim errorHandler As IErrorHandlerService
    Dim config2 As New CConfig
    Dim fileSystem2 As New CFileSystem
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config2, fileSystem2) ' Created here for top-level entry point
    
    ' Inicializar colección de suites
    Set m_SuiteNames = New Collection
    
    ' Descubrir y registrar automáticamente todas las suites disponibles
    DiscoverAndRegisterSuites
    
    ' Ejecutar todas las suites y obtener resultados
    Dim allResults As Collection
    Set allResults = ExecuteAllSuites
    
    ' Generar y mostrar reporte
    Dim reporter As New CTestReporter
    Call reporter.Initialize(allResults) ' Initialize reporter with results
    Dim reportString As String
    reportString = reporter.GenerateReport()
    MsgBox reportString, vbInformation, "Resultados de Pruebas CONDOR"
    
    Exit Sub
    
ErrorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.RunTestFramework", True ' Mark as critical
End Sub


'******************************************************************************
' GESTIÓN DE DESCUBRIMIENTO AUTOMÁTICO DE SUITES
'******************************************************************************

' Función que descubre automáticamente todas las suites de prueba basándose en convenciones de nomenclatura
' Convención: Los módulos de prueba deben comenzar con "Test" o "IntegrationTest"
' Patrón de función: [NombreModulo]RunAll (ej: TestCConfigRunAll, IntegrationTestCMapeoRepositoryRunAll)
' Requiere: Referencia a "Microsoft Visual Basic for Applications Extensibility 5.3"
Private Sub DiscoverAndRegisterSuites()
    On Error GoTo ErrorHandler
    
    ' Intentar descubrimiento automático primero
    Dim vbProject As Object
    Set vbProject = Application.VBE.ActiveVBProject
    
    ' Iterar sobre todos los componentes del proyecto
    Dim vbComponent As Object
    For Each vbComponent In vbProject.VBComponents
        ' Verificar si es un módulo estándar (Type = 1) y cumple con la convención de nomenclatura
        If vbComponent.Type = 1 Then ' vbext_ct_StdModule = 1
            Dim componentName As String
            componentName = vbComponent.Name
            
            ' Verificar si el nombre comienza con "Test" o "IntegrationTest"
            If Left(componentName, 4) = "Test" Or Left(componentName, 15) = "IntegrationTest" Then
                ' Construir el nombre de la función de ejecución siguiendo el patrón [NombreModulo]RunAll
                Dim suiteFunction As String
                suiteFunction = componentName & "RunAll"
                
                ' Añadir a la colección de suites
                m_SuiteNames.Add suiteFunction
            End If
        End If
    Next vbComponent
    
    Exit Sub
    
ErrorHandler:
    ' En caso de error en el descubrimiento automático, usar registro manual como fallback
    Dim errorHandler As IErrorHandlerService
    Dim config3 As New CConfig
    Dim fileSystem3 As New CFileSystem
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config3, fileSystem3)
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.DiscoverAndRegisterSuites - Usando fallback manual"
    
    ' Fallback: Registro manual de suites conocidas
    RegisterKnownSuites
End Sub

' Función de fallback que registra manualmente las suites conocidas
Private Sub RegisterKnownSuites()
    On Error Resume Next
    
    ' Registrar suites unitarias
    m_SuiteNames.Add "TestAppManagerRunAll"
    m_SuiteNames.Add "TestAuthServiceRunAll"
    m_SuiteNames.Add "TestCExpedienteServiceRunAll"
    m_SuiteNames.Add "TestCWordManagerRunAll"
    m_SuiteNames.Add "TestDocumentServiceRunAll"
    m_SuiteNames.Add "TestErrorHandlerServiceRunAll"
    m_SuiteNames.Add "TestModAssertRunAll"
    m_SuiteNames.Add "TestNotificationServiceRunAll"
    m_SuiteNames.Add "TestOperationLoggerRunAll"
    m_SuiteNames.Add "TestSolicitudServiceRunAll"
    m_SuiteNames.Add "TestWorkflowServiceRunAll"
    
    ' Registrar suites de integración
    m_SuiteNames.Add "IntegrationTestCConfigRunAll"
    m_SuiteNames.Add "IntegrationTestCExpedienteRepositoryRunAll"
    m_SuiteNames.Add "IntegrationTestCMapeoRepositoryRunAll"
    m_SuiteNames.Add "IntegrationTestNotificationServiceRunAll"
    m_SuiteNames.Add "IntegrationTestSolicitudRepositoryRunAll"
    m_SuiteNames.Add "IntegrationTestWordManagerRunAll"
    m_SuiteNames.Add "IntegrationTestWorkflowRepositoryRunAll"
    
    On Error GoTo 0
End Sub


'******************************************************************************
' MOTOR DE EJECUCIÓN
'******************************************************************************

' Función que ejecuta todas las suites registradas y devuelve resultados
Private Function ExecuteAllSuites() As Collection
    Dim allResults As New Collection
    Dim i As Integer
    
    For i = 1 To m_SuiteNames.count
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
            Call errorSuite.Initialize(suiteName)
            
            Dim errorTest As New CTestResult
            Call errorTest.Initialize("Suite_Execution_Error")
            Call errorTest.Fail("Error ejecutando suite: " & Err.Description)
            
            Call errorSuite.AddTest(errorTest)
            allResults.Add errorSuite
            
            ' Log the error
            Dim localErrorHandler As IErrorHandlerService
            Dim config4 As New CConfig
            Dim fileSystem4 As New CFileSystem
            Set localErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config4, fileSystem4)
            localErrorHandler.LogError Err.Number, Err.Description, "modTestRunner.ExecuteAllSuites", True ' Mark as critical
        End If
        
        On Error GoTo 0
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












