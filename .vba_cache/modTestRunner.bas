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


' Colección privada para registrar nombres de funciones de suite
Private m_SuiteNames As Scripting.Dictionary

' Función para compatibilidad con CLI (debe estar fuera del bloque condicional)
Public Function RunAllTests() As String
    On Error GoTo ErrorHandler
    
    ' 1. Crear una configuración específica para esta ejecución de pruebas
    Dim testConfig As New CMockConfig
    testConfig.SetSetting "LOG_FILE_PATH", CurrentProject.Path & "\condor_test_run.log"
    
    ' 2. Crear el ErrorHandler para el PROPIO RUNNER, inyectándole la config de prueba
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    
    ' El resto del flujo continúa...
    Set m_SuiteNames = New Scripting.Dictionary
    m_SuiteNames.CompareMode = TextCompare
    
    DiscoverAndRegisterSuites
    
    Dim reporter As ITestReporter
    Dim reporterImpl As New CTestReporter
    Set reporter = reporterImpl
    
    Dim allResults As Scripting.Dictionary
    Set allResults = ExecuteAllSuites ' ExecuteAllSuites usará la testConfig para sus propios error handlers
    
    reporter.Initialize allResults
    
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    RunAllTests = reportString
    Exit Function
    
ErrorHandler:
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modTestRunner.RunAllTests", True
    End If
    RunAllTests = "FALLO CRÍTICO EN EL MOTOR DE PRUEBAS: " & Err.Description & vbCrLf & "RESULT: FAILED"
End Function


' Alias para compatibilidad con CLI
Public Function ExecuteAllTests() As String
    ExecuteAllTests = RunAllTests()
End Function

' Función específica para CLI - Sin MsgBox, manejo robusto de errores
Public Function ExecuteAllTestsForCLI() As String
    On Error GoTo ErrorHandler
    
    ' 1. Crear una configuración específica para esta ejecución de pruebas
    Dim testConfig As New CMockConfig
    testConfig.SetSetting "LOG_FILE_PATH", CurrentProject.Path & "\condor_test_run.log"
    
    ' 2. Crear el ErrorHandler para el PROPIO RUNNER, inyectándole la config de prueba
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)

    ' El resto del flujo continúa...
    Set m_SuiteNames = New Scripting.Dictionary
    m_SuiteNames.CompareMode = TextCompare

    DiscoverAndRegisterSuites

    Dim reporter As ITestReporter
    Dim reporterImpl As New CTestReporter
    Set reporter = reporterImpl

    Dim allResults As Scripting.Dictionary
    Set allResults = ExecuteAllSuites ' ExecuteAllSuites usará la testConfig para sus propios error handlers

    reporter.Initialize allResults

    Dim reportString As String
    reportString = reporter.GenerateReport()

    ' Verificar si todas las pruebas pasaron
    Dim allPassed As Boolean
    allPassed = True

    Dim suiteResult As CTestSuiteResult
    Dim key As Variant
    For Each key In allResults.Keys()
        Set suiteResult = allResults(key)
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
    On Error GoTo errorHandler
    
    ' Obtener instancia del manejador de errores
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Inicializar colección de suites
    Set m_SuiteNames = New Scripting.Dictionary
    m_SuiteNames.CompareMode = TextCompare
    
    ' Descubrir y registrar automáticamente todas las suites disponibles
    DiscoverAndRegisterSuites
    
    ' 1. EJECUTAR
    Dim allResults As Scripting.Dictionary
    Set allResults = ExecuteAllSuites
    
    ' 2. GENERAR REPORTE (Responsabilidad de CTestReporter)
    Dim reporter As New CTestReporter
    reporter.Initialize allResults
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    ' 3. PRESENTAR (Responsabilidad del Runner/UI)
    MsgBox reportString, vbInformation, "Resultados de Pruebas CONDOR"
    
    Exit Sub
    
errorHandler:
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.RunTestFramework", True
End Sub


'******************************************************************************
' GESTIÓN DE DESCUBRIMIENTO AUTOMÁTICO DE SUITES
'******************************************************************************

' Función que descubre automáticamente todas las suites de prueba basándose en convenciones de nomenclatura
' Convención: Los módulos de prueba deben comenzar con "Test" o "IntegrationTest"
' Patrón de función: [NombreModulo]RunAll (ej: TestCConfigRunAll, IntegrationTestCMapeoRepositoryRunAll)
' Requiere: Referencia a "Microsoft Visual Basic for Applications Extensibility 5.3"
Private Sub DiscoverAndRegisterSuites()
    On Error GoTo errorHandler
    
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
            
            ' Verificar si el nombre comienza con "Test" o "TI"
            If LCase(Left(componentName, 4)) = "test" Or LCase(Left(componentName, 2)) = "ti" Then
                ' Construir el nombre de la función de ejecución siguiendo el patrón [NombreModulo]RunAll
                Dim suiteFunction As String
                suiteFunction = componentName & "RunAll"
                
                ' Añadir a la colección de suites
                m_SuiteNames.Add suiteFunction, suiteFunction
            End If
        End If
    Next vbComponent
    
    Exit Sub
    
errorHandler:
    ' El fallo del descubrimiento automático es ahora inaceptable - registrar error crítico
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.DiscoverAndRegisterSuites - FALLO CRÍTICO EN DESCUBRIMIENTO AUTOMÁTICO", True
    
    ' Re-lanzar el error ya que el descubrimiento automático debe funcionar
    Err.Raise Err.Number, "modTestRunner.DiscoverAndRegisterSuites", "FALLO CRÍTICO: " & Err.Description
End Sub




'******************************************************************************
' MOTOR DE EJECUCIÓN
'******************************************************************************

' Función que ejecuta todas las suites registradas y devuelve resultados
Private Function ExecuteAllSuites() As Scripting.Dictionary
    Dim allResults As New Scripting.Dictionary
    allResults.CompareMode = TextCompare
    Dim i As Integer
    
    Dim suiteKeys As Variant
    suiteKeys = m_SuiteNames.Keys()
    
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
            ' Crear un resultado de error para suites que fallan
            Dim errorSuite As New CTestSuiteResult
            Call errorSuite.Initialize(suiteName)
            
            Dim errorTest As New CTestResult
            Call errorTest.Initialize("Suite_Execution_Error")
            Call errorTest.Fail("Error ejecutando suite: " & Err.Description)
            
            Call errorSuite.AddResult(errorTest)
            allResults.Add suiteName, errorSuite
            
            ' Log the error
            Dim localErrorHandler As IErrorHandlerService
            Set localErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
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














