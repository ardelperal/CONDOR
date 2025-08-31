Attribute VB_Name = "TIExpedienteRepository"
Option Compare Database
Option Explicit

' =====================================================
' MODULO: IntegrationTestCExpedienteRepository
' DESCRIPCION: Pruebas de integración para CExpedienteRepository con BD real
' =====================================================

' Constantes para el autoaprovisionamiento de bases de datos
Private Const EXPEDIENTES_TEMPLATE_PATH As String = "back\test_db\templates\Expedientes_test_template.accdb"
Private Const EXPEDIENTES_ACTIVE_PATH As String = "back\test_db\active\Expedientes_integration_test.accdb"

' Variables globales para las dependencias reales
Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService
Private m_Repository As IExpedienteRepository

' Función principal que ejecuta todas las pruebas de integración del CExpedienteRepository
Public Function TIExpedienteRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIExpedienteRepository"
    
    suiteResult.AddTestResult IntegrationTestObtenerExpedientePorIdSuccess()
    suiteResult.AddTestResult IntegrationTestObtenerExpedientePorIdNotFound()
    suiteResult.AddTestResult IntegrationTestObtenerExpedientePorNemotecnicoSuccess()
    suiteResult.AddTestResult IntegrationTestObtenerExpedientePorNemotecnicoNotFound()
    suiteResult.AddTestResult IntegrationTestObtenerExpedientesActivosParaSelectorSuccess()
    suiteResult.AddTestResult IntegrationTestObtenerExpedientesActivosParaSelectorEmptyResult()
    
    Set TIExpedienteRepositoryRunAll = suiteResult
End Function

' ============================================================================ 
' SETUP Y TEARDOWN
' ============================================================================ 

Private Sub Setup()
    On Error GoTo ErrorHandler
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & EXPEDIENTES_TEMPLATE_PATH, modTestUtils.GetProjectPath() & EXPEDIENTES_ACTIVE_PATH
    InitializeRealDependencies
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "IntegrationTestCExpedienteRepository.Setup", "Error en Setup: " & Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Set m_Repository = Nothing
    Set m_ErrorHandler = Nothing
    Set m_Config = Nothing
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testDbPath As String
    testDbPath = modTestUtils.GetProjectPath() & EXPEDIENTES_ACTIVE_PATH
    If fs.FileExists(testDbPath) Then
        fs.DeleteFile testDbPath
    End If
    Set fs = Nothing
End Sub

Private Sub InitializeRealDependencies()
    On Error GoTo ErrorHandler
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & EXPEDIENTES_ACTIVE_PATH, "EXPEDIENTES_DB_PATH"
    Set m_Config = modConfigFactory.CreateConfigServiceFromCollection(settings)
    
    Set m_ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(m_Config)
    Set m_Repository = modRepositoryFactory.CreateExpedienteRepository(m_Config, m_ErrorHandler)
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "IntegrationTestCExpedienteRepository.InitializeRealDependencies", "Error inicializando dependencias: " & Err.Description
End Sub

' ============================================================================ 
' PRUEBAS DE INTEGRACIÓN
' ============================================================================ 

Private Function IntegrationTestObtenerExpedientePorIdSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ObtenerExpedientePorId debe devolver un expediente existente"
    Dim rs As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Setup
    
    Set rs = m_Repository.ObtenerExpedientePorId(1)
    
    modAssert.AssertNotNull rs, "ObtenerExpedientePorId debe devolver un recordset válido"
    modAssert.AssertFalse rs.EOF, "El recordset no debe estar vacío para expediente existente"
    modAssert.AssertNotNull rs.Fields("NumeroExpediente").Value, "NumeroExpediente no debe ser nulo"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Teardown
    Set IntegrationTestObtenerExpedientePorIdSuccess = testResult
End Function

Private Function IntegrationTestObtenerExpedientePorIdNotFound() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ObtenerExpedientePorId debe manejar expedientes no encontrados"
    Dim rs As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Setup
    
    Set rs = m_Repository.ObtenerExpedientePorId(99999)
    
    modAssert.AssertNotNull rs, "El recordset debe ser válido aunque esté vacío"
    modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para expediente no encontrado"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Teardown
    Set IntegrationTestObtenerExpedientePorIdNotFound = testResult
End Function

Private Function IntegrationTestObtenerExpedientePorNemotecnicoSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ObtenerExpedientePorNemotecnico debe devolver un expediente existente"
    Dim rs As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Setup
    
    Set rs = m_Repository.ObtenerExpedientePorNemotecnico("EXP-2024-001")
    
    modAssert.AssertNotNull rs, "ObtenerExpedientePorNemotecnico debe devolver un recordset válido"
    modAssert.AssertFalse rs.EOF, "El recordset no debe estar vacío para nemotécnico existente"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Teardown
    Set IntegrationTestObtenerExpedientePorNemotecnicoSuccess = testResult
End Function

Private Function IntegrationTestObtenerExpedientePorNemotecnicoNotFound() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ObtenerExpedientePorNemotecnico debe manejar nemotécnicos no encontrados"
    Dim rs As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Setup
    
    Set rs = m_Repository.ObtenerExpedientePorNemotecnico("INEXISTENTE-999")
    
    modAssert.AssertNotNull rs, "El recordset debe ser válido aunque esté vacío"
    modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para nemotécnico no encontrado"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Teardown
    Set IntegrationTestObtenerExpedientePorNemotecnicoNotFound = testResult
End Function

Private Function IntegrationTestObtenerExpedientesActivosParaSelectorSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ObtenerExpedientesActivosParaSelector debe devolver expedientes activos"
    Dim rs As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Setup
    
    Set rs = m_Repository.ObtenerExpedientesActivosParaSelector()
    
    modAssert.AssertNotNull rs, "ObtenerExpedientesActivosParaSelector debe devolver un recordset válido"
    modAssert.AssertFalse rs.EOF, "El recordset no debe estar vacío si hay expedientes activos"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Teardown
    Set IntegrationTestObtenerExpedientesActivosParaSelectorSuccess = testResult
End Function

Private Function IntegrationTestObtenerExpedientesActivosParaSelectorEmptyResult() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ObtenerExpedientesActivosParaSelector debe manejar cuando no hay expedientes activos"
    Dim rs As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Setup
    
    ' Vaciar la tabla para forzar el caso de resultado vacío
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(m_Config.GetValue("EXPEDIENTES_DB_PATH"))
    db.Execute "DELETE FROM T_Expedientes"
    db.Close
    Set db = Nothing
    
    Set rs = m_Repository.ObtenerExpedientesActivosParaSelector()
    
    modAssert.AssertNotNull rs, "El recordset debe ser válido aunque esté vacío"
    modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío si no hay expedientes activos"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Teardown
    Set IntegrationTestObtenerExpedientesActivosParaSelectorEmptyResult = testResult
End Function