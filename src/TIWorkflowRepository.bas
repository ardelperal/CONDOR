Attribute VB_Name = "TIWorkflowRepository"
' Módulo: TIWorkflowRepository
' Propósito: Pruebas de integración para CWorkflowRepository

Option Compare Database
Option Explicit

Private Const TEST_DB_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const TEST_DB_ACTIVE_PATH As String = "back\test_db\active\CONDOR_workflow_test.accdb"
Private Const TEST_TIPO_SOLICITUD As String = "PC"
Private Const TEST_ROL As String = "CALIDAD"
Private Const ESTADO_BORRADOR As String = "BORRADOR"
Private Const ESTADO_REVISION As String = "EN_REVISION"

Public Function TIWorkflowRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWorkflowRepository"

    suiteResult.AddTestResult TestGetNextStatesIntegration()
    suiteResult.AddTestResult TestIsValidTransitionIntegration()

    Set TIWorkflowRepositoryRunAll = suiteResult
End Function

Private Sub Setup()
    On Error GoTo TestError
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & TEST_DB_TEMPLATE_PATH, modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH
    InsertTestData modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH
    Exit Sub
TestError:
    Err.Raise Err.Number, "TIWorkflowRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    If fs.FileExists(modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH) Then
        fs.DeleteFile modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH, True
    End If
End Sub

Private Sub InsertTestData(ByVal dbPath As String)
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(dbPath, False, False)

    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (1, 'BORRADOR')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (2, 'EN_REVISION')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (3, 'APROBADO')"
    db.Execute "INSERT INTO TbTransiciones (idEstadoOrigen, idEstadoDestino, RolRequerido, TipoSolicitud) VALUES (1, 2, 'CALIDAD', 'PC')"
    db.Execute "INSERT INTO TbTransiciones (idEstadoOrigen, idEstadoDestino, RolRequerido, TipoSolicitud) VALUES (2, 3, 'ADMIN', 'PC')"

    db.Close
    Set db = Nothing
End Sub

Private Function TestGetNextStatesIntegration() As CTestResult
    Dim repository As IWorkflowRepository
    Dim config As IConfig
    Dim errorHandler As IErrorHandlerService
    
    Set TestGetNextStatesIntegration = New CTestResult
    TestGetNextStatesIntegration.Initialize "GetNextStates debe devolver los estados siguientes correctos"
    On Error GoTo TestFail

    Call Setup
    
    ' Crear dependencias localmente
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set repository = modRepositoryFactory.CreateWorkflowRepository()

    Dim nextStates As Collection
    Set nextStates = repository.GetNextStates(ESTADO_BORRADOR, TEST_TIPO_SOLICITUD, TEST_ROL)

    modAssert.AssertNotNull nextStates, "La colección de siguientes estados no debe ser nula."
    modAssert.AssertEquals 1, nextStates.Count, "Debe haber exactamente un estado siguiente desde BORRADOR."
    modAssert.AssertEquals "EN_REVISION", nextStates(1), "El siguiente estado debe ser EN_REVISION."

    TestGetNextStatesIntegration.Pass
    GoTo Cleanup
TestFail:
    TestGetNextStatesIntegration.Fail Err.Description
Cleanup:
    Call Teardown
End Function

Private Function TestIsValidTransitionIntegration() As CTestResult
    Dim repository As IWorkflowRepository
    Dim config As IConfig
    Dim errorHandler As IErrorHandlerService
    
    Set TestIsValidTransitionIntegration = New CTestResult
    TestIsValidTransitionIntegration.Initialize "IsValidTransition debe validar transiciones correctas e incorrectas"
    On Error GoTo TestFail

    Call Setup
    
    ' Crear dependencias localmente
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set repository = modRepositoryFactory.CreateWorkflowRepository()

    modAssert.AssertTrue repository.IsValidTransition(TEST_TIPO_SOLICITUD, ESTADO_BORRADOR, ESTADO_REVISION), "BORRADOR -> EN_REVISION debe ser una transición válida."
    modAssert.AssertFalse repository.IsValidTransition(TEST_TIPO_SOLICITUD, ESTADO_BORRADOR, "APROBADO"), "BORRADOR -> APROBADO debe ser una transición inválida."

    TestIsValidTransitionIntegration.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransitionIntegration.Fail Err.Description
Cleanup:
    Call Teardown
End Function
