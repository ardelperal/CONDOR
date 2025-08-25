Attribute VB_Name = "modTestRunner"

' Colección privada para registrar nombres de funciones de suite
Private m_SuiteNames As Collection

' Función para compatibilidad con CLI (debe estar fuera del bloque condicional)
Public Function RunAllTests() As String
    On Error GoTo ErrorHandler
    
    ' Inicializar colección de suites
    Set m_SuiteNames = New Collection
    
    ' Descubrir y registrar automáticamente todas las suites disponibles
    DiscoverAndRegisterSuites
    
    ' Ejecutar todas las pruebas y devolver resultado como string
    Dim reporter As ITestReporter
    Set reporter = New CTestReporter
    
    Dim allResults As Collection
    Set allResults = ExecuteAllSuites()
    
    ' Inicializar el reportero con los resultados
    reporter.Initialize allResults
    
    ' Generar reporte en formato string
    Dim reportString As String
    reportString = reporter.GenerateReport()
    
    RunAllTests = reportString
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.RunAllTests"
    
    RunAllTests = "FALLO CRÍTICO EN EL MOTOR DE PRUEBAS: " & Err.Description & vbCrLf & "RESULT: FAILED"
End Function

' Alias para compatibilidad con CLI
Public Function ExecuteAllTests() As String
    ExecuteAllTests = RunAllTests()
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
    
    ' Descubrir y registrar automáticamente todas las suites disponibles
    DiscoverAndRegisterSuites
    
    ' Ejecutar todas las suites y obtener resultados
    Dim allResults As Collection
    Set allResults = ExecuteAllSuites()
    
    ' Generar y mostrar reporte
    Dim reporter As New CTestReporter
    reporter.ShowReport allResults
End Sub

'******************************************************************************
' GESTIÓN DE DESCUBRIMIENTO AUTOMÁTICO DE SUITES
'******************************************************************************

' Función que descubre automáticamente todas las suites de prueba basándose en convenciones de nomenclatura
' Convención: Los módulos de prueba deben comenzar con "Test_" o "IntegrationTest_"
' Patrón de función: [NombreModulo]_RunAll (ej: Test_CConfig_RunAll, IntegrationTest_CMapeoRepository_RunAll)
' Requiere: Referencia a "Microsoft Visual Basic for Applications Extensibility 5.3"
Private Sub DiscoverAndRegisterSuites()
    On Error GoTo ErrorHandler
    
    ' Obtener referencia al proyecto VBA actual
    Dim vbProject As Object
    Set vbProject = Application.VBE.ActiveVBProject
    
    ' Iterar sobre todos los componentes del proyecto
    Dim vbComponent As Object
    For Each vbComponent In vbProject.VBComponents
        ' Verificar si es un módulo estándar (Type = 1) y cumple con la convención de nomenclatura
        If vbComponent.Type = 1 Then ' vbext_ct_StdModule = 1
            Dim componentName As String
            componentName = vbComponent.Name
            
            ' Verificar si el nombre comienza con "Test_" o "IntegrationTest_"
            If Left(componentName, 5) = "Test_" Or Left(componentName, 16) = "IntegrationTest_" Then
                ' Construir el nombre de la función de ejecución siguiendo el patrón [NombreModulo]_RunAll
                Dim suiteFunction As String
                suiteFunction = componentName & "_RunAll"
                
                ' Añadir a la colección de suites
                m_SuiteNames.Add suiteFunction
            End If
        End If
    Next vbComponent
    
    Exit Sub
    
ErrorHandler:
    ' En caso de error en el descubrimiento, registrar el error pero continuar
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modTestRunner.DiscoverAndRegisterSuites"
    
    ' Si falla el descubrimiento automático, al menos intentar registrar las suites conocidas críticas
    On Error Resume Next
    m_SuiteNames.Add "Test_CConfig_RunAll"
    m_SuiteNames.Add "Test_AuthService_RunAll"
    On Error GoTo 0
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









