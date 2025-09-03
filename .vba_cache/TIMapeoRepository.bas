Attribute VB_Name = "TIMapeoRepository"
Option Compare Database
Option Explicit


Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\active\CONDOR_mapeo_integration_test.accdb"

Public Function TIMapeoRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIMapeoRepository"
    suiteResult.AddResult TestGetMapeoPorTipoSuccess()
    suiteResult.AddResult TestGetMapeoPorTipoNotFound()
    Set TIMapeoRepositoryRunAll = suiteResult
End Function

Private Sub Setup()
    On Error GoTo TestError
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & CONDOR_TEMPLATE_PATH, modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH)
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'refContrato', 'MARCADOR_CONTRATO')"
    db.Close
    Set db = Nothing
    Exit Sub
TestError:
    Err.Raise Err.Number, "TIMapeoRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFile modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, True
    Set fs = Nothing
End Sub

Private Function TestGetMapeoPorTipoSuccess() As CTestResult
    Set TestGetMapeoPorTipoSuccess = New CTestResult
    TestGetMapeoPorTipoSuccess.Initialize "GetMapeoPorTipo debe devolver un objeto EMapeo con datos"
    Dim repository As IMapeoRepository, config As IConfig, errorHandler As IErrorHandlerService, mapeoResult As EMapeo
    On Error GoTo TestFail
    Call Setup

    ' ARRANGE: Crear y poblar COMPLETAMENTE la configuración de prueba
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    config.SetSetting "DB_PASSWORD", ""
    config.SetSetting "LOG_FILE_PATH", modTestUtils.GetProjectPath() & "back\test_db\active\test_run.log"

    ' Crear el resto de las dependencias usando la configuración ya poblada
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config)
    Set repository = modRepositoryFactory.CreateMapeoRepository(config, errorHandler)
    Set mapeoResult = repository.GetMapeoPorTipo("PC")
    modAssert.AssertNotNull mapeoResult, "El objeto EMapeo no debe ser nulo."
    AssertEquals "PC", mapeoResult.NombrePlantilla, "El nombre de la plantilla no es el esperado."
    TestGetMapeoPorTipoSuccess.Pass
    GoTo Cleanup
TestFail:
    TestGetMapeoPorTipoSuccess.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set repository = Nothing: Set config = Nothing: Set errorHandler = Nothing: Set mapeoResult = Nothing
    Call Teardown
End Function

Private Function TestGetMapeoPorTipoNotFound() As CTestResult
    Set TestGetMapeoPorTipoNotFound = New CTestResult
    TestGetMapeoPorTipoNotFound.Initialize "GetMapeoPorTipo debe devolver Nothing si no hay mapeo"
    Dim repository As IMapeoRepository, config As IConfig, errorHandler As IErrorHandlerService, mapeoResult As EMapeo
    On Error GoTo TestFail
    Call Setup

    ' ARRANGE: Crear y poblar COMPLETAMENTE la configuración de prueba
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    config.SetSetting "DB_PASSWORD", ""
    config.SetSetting "LOG_FILE_PATH", modTestUtils.GetProjectPath() & "back\test_db\active\test_run.log"

    ' Crear el resto de las dependencias usando la configuración ya poblada
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config)
    Set repository = modRepositoryFactory.CreateMapeoRepository(config, errorHandler)
    Set mapeoResult = repository.GetMapeoPorTipo("TIPO_INEXISTENTE")
    AssertIsNull mapeoResult, "El objeto EMapeo devuelto debería ser Nothing."
    TestGetMapeoPorTipoNotFound.Pass
    GoTo Cleanup
TestFail:
    TestGetMapeoPorTipoNotFound.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set repository = Nothing: Set config = Nothing: Set errorHandler = Nothing: Set mapeoResult = Nothing
    Call Teardown
End Function

