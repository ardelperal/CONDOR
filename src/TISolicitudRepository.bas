Attribute VB_Name = "TISolicitudRepository"
Option Compare Database
Option Explicit

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

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
    On Error GoTo ErrorHandler
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    
    ' Usar las constantes ya definidas para construir los nombres de archivo
    Dim templateDbName As String: templateDbName = "Solicitud_test_template.accdb"
    Dim activeDbName As String: activeDbName = "Solicitud_integration_test.accdb"
    
    ' Llamada al método correcto de modTestUtils
    modTestUtils.PrepareTestDatabase templateDbName, activeDbName
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TISolicitudRepository.SuiteSetup", Err.Description
End Sub

Private Sub SuiteTeardown()
    ' Limpieza estandarizada a través de la utilidad central.
    Call modTestUtils.CleanupTestDatabase("Solicitud_integration_test.accdb")
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
    Dim fs As IFileSystem
    Dim dbPath As String
    
    On Error GoTo TestFail

    ' Arrange: Crear configuración local apuntando a la BD de prueba de esta suite
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "CONDOR_DATA_PATH", modTestUtils.GetWorkspacePath() & "Solicitud_integration_test.accdb"
    mockConfigImpl.SetSetting "CONDOR_PASSWORD", ""
    Set localConfig = mockConfigImpl
    
    ' Arrange: Crear dependencias inyectando la configuración local
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set repo = modRepositoryFactory.CreateSolicitudRepository(localConfig)
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = modTestUtils.GetWorkspacePath() & "Solicitud_integration_test.accdb"
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 102, "Test.Arrange", "La BD de prueba de Solicitudes no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath, False, False)
    
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
    Set fs = Nothing
End Function

