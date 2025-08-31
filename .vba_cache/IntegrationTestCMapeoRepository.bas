Attribute VB_Name = "IntegrationTestCMapeoRepository"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS DE INTEGRACIÓN PARA CMapeoRepository
' Arquitectura: Autoaprovisionamiento con Base de Datos Real
' ============================================================================

' Constantes para rutas de base de datos de prueba
Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\active\CONDOR_mapeo_integration_test.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function IntegrationTestCMapeoRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("IntegrationTestCMapeoRepository")
    
    Call suiteResult.AddTestResult(TestGetMapeoPorTipoSuccess())
    Call suiteResult.AddTestResult(TestGetMapeoPorTipoNotFound())
    
    Set IntegrationTestCMapeoRepositoryRunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    On Error GoTo TestError
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & CONDOR_TEMPLATE_PATH, modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Exit Sub
TestError:
    Err.Raise Err.Number, "IntegrationTestCMapeoRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testPath As String
    testPath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    If fs.FileExists(testPath) Then
        fs.DeleteFile testPath
    End If
    Set fs = Nothing
End Sub

' ============================================================================
' PRUEBAS
' ============================================================================

Private Function TestGetMapeoPorTipoSuccess() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetMapeoPorTipo debe devolver un recordset con datos")
    
    ' Variables locales para dependencias
    Dim repository As IMapeoRepository
    Dim config As IConfig
    Dim errorHandler As IErrorHandlerService
    
    On Error GoTo TestFail
    
    Call Setup
    
    ' Configurar dependencias con la base de datos de prueba
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    
    Set config = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, modFileSystemFactory.CreateFileSystem())
    Set repository = modRepositoryFactory.CreateMapeoRepository(config, errorHandler)
    
    Dim tipoSolicitud As String
    tipoSolicitud = "PC"
    
    Dim rs As DAO.recordset
    Set rs = repository.GetMapeoPorTipo(tipoSolicitud)
    
    AssertNotNull rs, "El recordset no debe ser nulo"
    AssertTrue Not rs.EOF, "El recordset no debe estar vacío"
    
    rs.Close
    Set rs = Nothing
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    Set repository = Nothing
    Set config = Nothing
    Set errorHandler = Nothing
    Call Teardown
    Set TestGetMapeoPorTipoSuccess = testResult
End Function

Private Function TestGetMapeoPorTipoNotFound() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetMapeoPorTipo debe devolver un recordset vacío si no hay mapeo")
    
    ' Variables locales para dependencias
    Dim repository As IMapeoRepository
    Dim config As IConfig
    Dim errorHandler As IErrorHandlerService
    
    On Error GoTo TestFail
    
    Call Setup
    
    ' Configurar dependencias con la base de datos de prueba
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    
    Set config = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, modFileSystemFactory.CreateFileSystem())
    Set repository = modRepositoryFactory.CreateMapeoRepository(config, errorHandler)
    
    Dim tipoSolicitud As String
    tipoSolicitud = "TIPO_INEXISTENTE"
    
    Dim rs As DAO.recordset
    Set rs = repository.GetMapeoPorTipo(tipoSolicitud)
    
    AssertNotNull rs, "El recordset no debe ser nulo"
    AssertTrue rs.EOF, "El recordset debe estar vacío"
    
    rs.Close
    Set rs = Nothing
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    Set repository = Nothing
    Set config = Nothing
    Set errorHandler = Nothing
    Call Teardown
    Set TestGetMapeoPorTipoNotFound = testResult
End Function

' ============================================================================
' FUNCIONES DE AUTOAPROVISIONAMIENTO
' ============================================================================