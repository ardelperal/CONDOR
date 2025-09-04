Attribute VB_Name = "TISolicitudRepository"
Option Compare Database
Option Explicit

Private Const TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const ACTIVE_PATH As String = "back\test_db\active\CONDOR_solicitud_itest.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TISolicitudRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TISolicitudRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestSaveAndRetrieveSolicitud()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TISolicitudRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & TEMPLATE_PATH
    Dim activePath As String: activePath = projectPath & ACTIVE_PATH
    Call modTestUtils.SuiteSetup(templatePath, activePath)
End Sub

Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & ACTIVE_PATH
    Call modTestUtils.SuiteTeardown(activePath)
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function TestSaveAndRetrieveSolicitud() As CTestResult
    Set TestSaveAndRetrieveSolicitud = New CTestResult
    TestSaveAndRetrieveSolicitud.Initialize "Debe guardar y recuperar una solicitud correctamente"
    
    Dim localConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim repo As ISolicitudRepository
    Dim retrievedSolicitud As ESolicitud
    Dim db As DAO.Database
    
    On Error GoTo TestFail

    ' Arrange: 1. Crear configuración local apuntando a la BD de prueba activa
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", modTestUtils.GetProjectPath() & ACTIVE_PATH
    Set localConfig = mockConfigImpl
    
    ' Arrange: 2. Crear dependencias inyectando la configuración local
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(localConfig)
    Set repo = modRepositoryFactory.CreateSolicitudRepository(localConfig, errorHandler)
    Set db = DBEngine.OpenDatabase(localConfig.GetDataPath(), False, False)
    
    Dim nuevaSolicitud As New ESolicitud
    nuevaSolicitud.idExpediente = 999
    nuevaSolicitud.tipoSolicitud = "TIPO_TEST"
    nuevaSolicitud.codigoSolicitud = "TEST-SAVE-001"
    nuevaSolicitud.idEstadoInterno = 1 ' Borrador
    nuevaSolicitud.usuarioCreacion = "itest_user"
    
    ' Act: Guardar y recuperar dentro de una transacción
    DBEngine.BeginTrans
    Dim newId As Long
    newId = repo.SaveSolicitud(nuevaSolicitud)
    Set retrievedSolicitud = repo.ObtenerSolicitudPorId(newId)
    
    ' Assert
    modAssert.AssertTrue newId > 0, "El ID devuelto debe ser positivo."
    modAssert.AssertNotNull retrievedSolicitud, "La solicitud recuperada no debe ser nula."
    modAssert.AssertEquals "TEST-SAVE-001", retrievedSolicitud.codigoSolicitud, "El código no coincide."
    modAssert.AssertEquals "itest_user", retrievedSolicitud.usuarioCreacion, "El usuario no coincide."

    TestSaveAndRetrieveSolicitud.Pass
    GoTo Cleanup

TestFail:
    TestSaveAndRetrieveSolicitud.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set retrievedSolicitud = Nothing
    Set repo = Nothing
    Set errorHandler = Nothing
    Set localConfig = Nothing
    Set db = Nothing
End Function

