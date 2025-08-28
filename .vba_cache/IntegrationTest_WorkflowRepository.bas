Attribute VB_Name = "IntegrationTest_WorkflowRepository"
' Módulo: IntegrationTest_WorkflowRepository
' Propósito: Pruebas de integración para CWorkflowRepository
' Versión: 2.0 (Reescritura completa)

Option Compare Database
Option Explicit

#If DEV_MODE Then

Private Const TEST_DB_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const TEST_DB_ACTIVE_PATH As String = "back\test_db\active\CONDOR_workflow_test.accdb"
Private Const TEST_TIPO_SOLICITUD As String = "PC"
Private Const TEST_ROL As String = "CALIDAD"
Private Const ESTADO_BORRADOR As String = "BORRADOR"
Private Const ESTADO_REVISION As String = "EN_REVISION"

Private m_Repository As IWorkflowRepository
Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Function IntegrationTest_WorkflowRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_WorkflowRepository"

    suiteResult.AddTestResult Test_GetNextStates_Integration()
    suiteResult.AddTestResult Test_IsValidTransition_Integration()
    ' ... aquí se pueden añadir el resto de tests ...

    Set IntegrationTest_WorkflowRepository_RunAll = suiteResult
End Function

Private Sub Setup()
    On Error GoTo TestError
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & TEST_DB_TEMPLATE_PATH, modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH

    Dim tempConfig As New CConfig
    tempConfig.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH
    Set m_Config = tempConfig

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Set m_ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(m_Config, fs)

    Set m_Repository = modRepositoryFactory.CreateWorkflowRepository(m_Config, m_ErrorHandler)

    InsertTestData
    Exit Sub
TestError:
    Err.Raise Err.Number, "IntegrationTest_WorkflowRepository.Setup", Err.Description
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

    ' Insertar estados de prueba con IDs explícitos
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (1, 'BORRADOR')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (2, 'EN_REVISION')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado) VALUES (3, 'APROBADO')"

    ' Insertar transiciones usando los IDs correctos
    db.Execute "INSERT INTO TbTransiciones (idEstadoOrigen, idEstadoDestino, RolRequerido, TipoSolicitud) VALUES (1, 2, 'CALIDAD', 'PC')"
    db.Execute "INSERT INTO TbTransiciones (idEstadoOrigen, idEstadoDestino, RolRequerido, TipoSolicitud) VALUES (2, 3, 'ADMIN', 'PC')"

    db.Close
    Set db = Nothing
End Sub

Private Function Test_GetNextStates_Integration() As CTestResult
    Set Test_GetNextStates_Integration = New CTestResult
    Test_GetNextStates_Integration.Initialize "GetNextStates debe devolver los estados siguientes correctos"
    On Error GoTo TestFail

    Call Setup

    ' Act
    Dim nextStates As Collection
    Set nextStates = m_Repository.GetNextStates(ESTADO_BORRADOR, TEST_TIPO_SOLICITUD, TEST_ROL)

    ' Assert
    modAssert.AssertNotNull nextStates, "La colección de siguientes estados no debe ser nula."
    modAssert.AssertEquals 1, nextStates.Count, "Debe haber exactamente un estado siguiente desde BORRADOR."
    modAssert.AssertEquals "EN_REVISION", nextStates(1), "El siguiente estado debe ser EN_REVISION."

    Test_GetNextStates_Integration.Pass
    GoTo Cleanup
TestFail:
    Test_GetNextStates_Integration.Fail Err.Description
Cleanup:
    Call Teardown
End Function

Private Function Test_IsValidTransition_Integration() As CTestResult
    Set Test_IsValidTransition_Integration = New CTestResult
    Test_IsValidTransition_Integration.Initialize "IsValidTransition debe validar transiciones correctas e incorrectas"
    On Error GoTo TestFail

    Call Setup

    ' Act & Assert
    modAssert.AssertTrue m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, ESTADO_BORRADOR, ESTADO_REVISION), "BORRADOR -> EN_REVISION debe ser una transición válida."
    modAssert.AssertFalse m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, ESTADO_BORRADOR, "APROBADO"), "BORRADOR -> APROBADO debe ser una transición inválida."

    Test_IsValidTransition_Integration.Pass
    GoTo Cleanup
TestFail:
    Test_IsValidTransition_Integration.Fail Err.Description
Cleanup:
    Call Teardown
End Function

#End If