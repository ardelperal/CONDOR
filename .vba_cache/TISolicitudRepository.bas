Attribute VB_Name = "TISolicitudRepository"
Option Compare Database
Option Explicit


Private Const TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const ACTIVE_PATH As String = "back\test_db\active\CONDOR_solicitud_itest.accdb"

Public Function TISolicitudRepositoryRunAll() As CTestSuiteResult
    Set TISolicitudRepositoryRunAll = New CTestSuiteResult
    TISolicitudRepositoryRunAll.Initialize "TISolicitudRepository"
    TISolicitudRepositoryRunAll.AddResult TestSaveAndRetrieveSolicitud()
End Function

Private Sub Setup()
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath & TEMPLATE_PATH, modTestUtils.GetProjectPath & ACTIVE_PATH
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFile modTestUtils.GetProjectPath & ACTIVE_PATH, True
End Sub

Private Function TestSaveAndRetrieveSolicitud() As CTestResult
    Set TestSaveAndRetrieveSolicitud = New CTestResult
    TestSaveAndRetrieveSolicitud.Initialize "Debe guardar y recuperar una solicitud"
    
    Dim repo As ISolicitudRepository, config As IConfig, newId As Long, retrievedSolicitud As ESolicitud
    On Error GoTo TestFail
    Call Setup
    
    ' Arrange
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath & ACTIVE_PATH
    config.SetSetting "LOG_FILE_PATH", modTestUtils.GetProjectPath() & "back\test_db\active\test_run.log"
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config)
    
    Set repo = modRepositoryFactory.CreateSolicitudRepository(config, errorHandler)
    
    Dim nuevaSolicitud As New ESolicitud
    nuevaSolicitud.idExpediente = 999
    nuevaSolicitud.tipoSolicitud = "TIPO_TEST"
    nuevaSolicitud.subTipoSolicitud = "SUBTIPO_TEST"
    nuevaSolicitud.codigoSolicitud = "TEST-SAVE-001"
    nuevaSolicitud.idEstadoInterno = 1 ' Estado Borrador
    nuevaSolicitud.usuarioCreacion = "integration_test_user"
    
    ' Act: Guardar y recuperar
    newId = repo.SaveSolicitud(nuevaSolicitud)
    Set retrievedSolicitud = repo.ObtenerSolicitudPorId(newId)
    
    ' Assert
    modAssert.AssertTrue newId > 0, "El ID devuelto debe ser positivo."
    modAssert.AssertNotNull retrievedSolicitud, "La solicitud recuperada no debe ser nula."
    modAssert.AssertEquals "TEST-SAVE-001", retrievedSolicitud.codigoSolicitud, "El código de solicitud no coincide."
    modAssert.AssertEquals 1, retrievedSolicitud.idEstadoInterno, "El estado interno debe ser 1 (Borrador)."
    modAssert.AssertEquals "integration_test_user", retrievedSolicitud.usuarioCreacion, "El usuario de creación no coincide."
    
    TestSaveAndRetrieveSolicitud.Pass
Cleanup:
    Call Teardown
    Exit Function
TestFail:
    TestSaveAndRetrieveSolicitud.Fail "Error: " & Err.Description
    Resume Cleanup
End Function

