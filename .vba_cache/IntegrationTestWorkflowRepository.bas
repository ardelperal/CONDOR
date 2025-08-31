Attribute VB_Name = "IntegrationTestWorkflowRepository"
' Módulo: IntegrationTestWorkflowRepository
' Propósito: Pruebas de integración para CWorkflowRepository

Option Compare Database
Option Explicit

Private Const TEST_DB_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const TEST_DB_ACTIVE_PATH As String = "back\test_db\active\CONDOR_workflow_test.accdb"
Private Const TEST_TIPO_SOLICITUD As String = "PC"
Private Const TEST_ROL As String = "CALIDAD"
Private Const ESTADO_BORRADOR As String = "BORRADOR"
Private Const ESTADO_REVISION As String = "EN_REVISION"

Private m_Repository As IWorkflowRepository
Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Function IntegrationTestWorkflowRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTestWorkflowRepository"

    suiteResult.AddTestResult TestGetNextStatesIntegration()
    suiteResult.AddTestResult TestIsValidTransitionIntegration()

    Set IntegrationTestWorkflowRepositoryRunAll = suiteResult
End Function

Private Sub Setup()
    On Error GoTo TestError
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & TEST_DB_TEMPLATE_PATH, modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH

    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH, "DATABASE_PATH"
    Set m_Config = modConfigFactory.CreateConfigServiceFromCollection(settings)

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Set m_ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(m_Config, fs)

    Set m_Repository = modRepositoryFactory.CreateWorkflowRepository(m_Config, m_ErrorHandler)

    InsertTestData
    Exit Sub
TestError:
    Err.Raise Err.Number, "IntegrationTestWorkflowRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    If fs.FileExists(modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH) Then
        fs.DeleteFile modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH, True
    End If
    Set m_Repository = Nothing
    Set m_Config = Nothing
    Set m_ErrorHandler = Nothing
End Sub

Private Sub InsertTestData()
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(m_Config.GetValue("DATABASE_PATH"), False, False)

    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (1, 'BORRADOR')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (2, 'EN_REVISION')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (3, 'APROBADO')"
    db.Execute "INSERT INTO TbTransiciones (idEstadoOrigen, idEstadoDestino, RolRequerido, TipoSolicitud) VALUES (1, 2, 'CALIDAD', 'PC')"
    db.Execute "INSERT INTO TbTransiciones (idEstadoOrigen, idEstadoDestino, RolRequerido, TipoSolicitud) VALUES (2, 3, 'ADMIN', 'PC')"

    db.Close
    Set db = Nothing
End Sub

Private Function TestGetNextStatesIntegration() As CTestResult
    Set TestGetNextStatesIntegration = New CTestResult
    TestGetNextStatesIntegration.Initialize "GetNextStates debe devolver los estados siguientes correctos"
    On Error GoTo TestFail

    Call Setup

    Dim nextStates As Collection
    Set nextStates = m_Repository.GetNextStates(ESTADO_BORRADOR, TEST_TIPO_SOLICITUD, TEST_ROL)

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
    Set TestIsValidTransitionIntegration = New CTestResult
    TestIsValidTransitionIntegration.Initialize "IsValidTransition debe validar transiciones correctas e incorrectas"
    On Error GoTo TestFail

    Call Setup

    modAssert.AssertTrue m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, ESTADO_BORRADOR, ESTADO_REVISION), "BORRADOR -> EN_REVISION debe ser una transición válida."
    modAssert.AssertFalse m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, ESTADO_BORRADOR, "APROBADO"), "BORRADOR -> APROBADO debe ser una transición inválida."

    TestIsValidTransitionIntegration.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransitionIntegration.Fail Err.Description
Cleanup:
    Call Teardown
End Function
