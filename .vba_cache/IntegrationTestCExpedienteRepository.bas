Attribute VB_Name = "IntegrationTest_CExpedienteRepository"
Option Compare Database
Option Explicit


' =====================================================
' MODULO: IntegrationTest_CExpedienteRepository
' DESCRIPCION: Pruebas de integración para CExpedienteRepository con BD real
' AUTOR: Sistema CONDOR
' FECHA: 2025
' =====================================================

#If DEV_MODE Then

' Constantes para el autoaprovisionamiento de bases de datos
Private Const EXPEDIENTES_TEMPLATE_PATH As String = "back\test_db\templates\Expedientes_test_template.accdb"
Private Const EXPEDIENTES_ACTIVE_PATH As String = "back\test_db\active\Expedientes_integration_test.accdb"

' Variables globales para las dependencias reales
Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService
Private m_Repository As CExpedienteRepository

' Función principal que ejecuta todas las pruebas de integración del CExpedienteRepository
Public Function IntegrationTestCExpedienteRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.Initialize "IntegrationTestCExpedienteRepository"
    
    ' Ejecutar todas las pruebas individuales
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorId_Success()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorId_NotFound()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorNemotecnico_Success()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientesActivosParaSelector_Success()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult()
    
    Set IntegrationTestCExpedienteRepository_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

' Configura el entorno de prueba con base de datos real
Private Sub Setup()
    On Error GoTo ErrorHandler
    
    ' Aprovisionar la base de datos de prueba usando el sistema estándar
    Dim fullTemplatePath As String
    Dim fullTestPath As String
    
    fullTemplatePath = modTestUtils.GetProjectPath() & EXPEDIENTES_TEMPLATE_PATH
    fullTestPath = modTestUtils.GetProjectPath() & EXPEDIENTES_ACTIVE_PATH
    
    modTestUtils.PrepareTestDatabase fullTemplatePath, fullTestPath
    
    ' Crear dependencias reales
    InitializeRealDependencies
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "IntegrationTest_CExpedienteRepository.Setup", "Error en Setup: " & Err.Description
End Sub

' Limpia el entorno de prueba
Private Sub Teardown()
    On Error Resume Next
    
    ' Limpiar referencias
    Set m_Repository = Nothing
    Set m_ErrorHandler = Nothing
    Set m_Config = Nothing
    
    ' Eliminar BD de prueba usando IFileSystem
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim testDbPath As String
    testDbPath = modTestUtils.GetProjectPath() & EXPEDIENTES_ACTIVE_PATH
    
    If fs.FileExists(testDbPath) Then
        fs.DeleteFile testDbPath
    End If
    
    Set fs = Nothing
End Sub

' Inicializa las dependencias reales para las pruebas
Private Sub InitializeRealDependencies()
    On Error GoTo ErrorHandler
    
    ' Crear config real que apunte a la BD de prueba
    Set m_Config = New CConfig
    
    ' Sobrescribir la ruta de BD para apuntar a la BD de prueba
    m_Config.SetSetting "EXPEDIENTES_DB_PATH", modTestUtils.GetProjectPath() & EXPEDIENTES_ACTIVE_PATH
    
    ' Crear errorHandler real
    Set m_ErrorHandler = New CErrorHandlerService
    
    ' Crear repositorio real usando factory
    Set m_Repository = modRepositoryFactory.CreateExpedienteRepository(m_Config, m_ErrorHandler)
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "IntegrationTest_CExpedienteRepository.InitializeRealDependencies", "Error inicializando dependencias: " & Err.Description
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN PARA ObtenerExpedientePorId
' ============================================================================

' Prueba que ObtenerExpedientePorId devuelve correctamente un expediente existente
Private Function IntegrationTest_ObtenerExpedientePorId_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorId_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar entorno de prueba real
    Setup
    
    ' Act - Ejecutar el método bajo prueba con ID conocido
    Dim rs As DAO.Recordset
    Set rs = m_Repository.ObtenerExpedientePorId(1)
    
    ' Assert - Verificar que devuelve un recordset válido con datos esperados
    modAssert.AssertNotNull rs, "ObtenerExpedientePorId debe devolver un recordset válido"
    modAssert.AssertFalse rs.EOF, "El recordset no debe estar vacío para expediente existente"
    
    ' Verificar campos específicos si existen
    If Not rs.EOF Then
        modAssert.AssertNotNull rs.Fields("NumeroExpediente").Value, "NumeroExpediente no debe ser nulo"
        modAssert.AssertNotNull rs.Fields("Titulo").Value, "Titulo no debe ser nulo"
    End If
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Teardown
    Set IntegrationTest_ObtenerExpedientePorId_Success = testResult
End Function

' Prueba que ObtenerExpedientePorId maneja correctamente expedientes no encontrados
Private Function IntegrationTest_ObtenerExpedientePorId_NotFound() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorId_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar entorno de prueba real
    Setup
    
    ' Act - Ejecutar el método con ID inexistente
    Dim rs As DAO.Recordset
    Set rs = m_Repository.ObtenerExpedientePorId(99999)
    
    ' Assert - Verificar que maneja correctamente el caso no encontrado
    modAssert.AssertNotNull rs, "El recordset debe ser válido aunque esté vacío"
    modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para expediente no encontrado"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Teardown
    Set IntegrationTest_ObtenerExpedientePorId_NotFound = testResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN PARA ObtenerExpedientePorNemotecnico
' ============================================================================

' Prueba que ObtenerExpedientePorNemotecnico devuelve correctamente un expediente existente
Private Function IntegrationTest_ObtenerExpedientePorNemotecnico_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorNemotecnico_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar entorno de prueba real
    Setup
    
    ' Act - Ejecutar el método bajo prueba con nemotécnico conocido
    Dim rs As DAO.Recordset
    Set rs = m_Repository.ObtenerExpedientePorNemotecnico("EXP-2024-001")
    
    ' Assert - Verificar que devuelve un recordset válido
    modAssert.AssertNotNull rs, "ObtenerExpedientePorNemotecnico debe devolver un recordset válido"
    
    ' Si encuentra el expediente, verificar campos
    If Not rs.EOF Then
        modAssert.AssertNotNull rs.Fields("NumeroExpediente").Value, "NumeroExpediente no debe ser nulo"
        modAssert.AssertNotNull rs.Fields("Titulo").Value, "Titulo no debe ser nulo"
    End If
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Teardown
    Set IntegrationTest_ObtenerExpedientePorNemotecnico_Success = testResult
End Function

' Prueba que ObtenerExpedientePorNemotecnico maneja correctamente nemotécnicos no encontrados
Private Function IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar entorno de prueba real
    Setup
    
    ' Act - Ejecutar el método con nemotécnico inexistente
    Dim rs As DAO.Recordset
    Set rs = m_Repository.ObtenerExpedientePorNemotecnico("INEXISTENTE-999")
    
    ' Assert - Verificar que maneja correctamente el caso no encontrado
    modAssert.AssertNotNull rs, "El recordset debe ser válido aunque esté vacío"
    modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para nemotécnico no encontrado"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Teardown
    Set IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound = testResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN PARA ObtenerExpedientesActivosParaSelector
' ============================================================================

' Prueba que ObtenerExpedientesActivosParaSelector devuelve correctamente expedientes activos
Private Function IntegrationTest_ObtenerExpedientesActivosParaSelector_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientesActivosParaSelector_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar entorno de prueba real
    Setup
    
    ' Act - Ejecutar el método bajo prueba
    Dim rs As DAO.Recordset
    Set rs = m_Repository.ObtenerExpedientesActivosParaSelector()
    
    ' Assert - Verificar que devuelve un recordset válido
    modAssert.AssertNotNull rs, "ObtenerExpedientesActivosParaSelector debe devolver un recordset válido"
    
    ' Si hay expedientes activos, verificar estructura
    If Not rs.EOF Then
        modAssert.AssertNotNull rs.Fields("NumeroExpediente").Value, "NumeroExpediente no debe ser nulo"
        modAssert.AssertNotNull rs.Fields("Titulo").Value, "Titulo no debe ser nulo"
    End If
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Teardown
    Set IntegrationTest_ObtenerExpedientesActivosParaSelector_Success = testResult
End Function

' Prueba que ObtenerExpedientesActivosParaSelector maneja correctamente cuando no hay expedientes activos
Private Function IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar entorno de prueba real
    Setup
    
    ' Act - Ejecutar el método bajo prueba
    Dim rs As DAO.Recordset
    Set rs = m_Repository.ObtenerExpedientesActivosParaSelector()
    
    ' Assert - Verificar que maneja correctamente el caso sin resultados
    modAssert.AssertNotNull rs, "El recordset debe ser válido aunque esté vacío"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Teardown
    Set IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult = testResult
End Function

#End If

