Attribute VB_Name = "TINotificationService"
Option Compare Database
Option Explicit

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TINotificationServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TINotificationService (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestSendNotificationSuccessCallsRepositoryCorrectly()
    suiteResult.AddResult TestInitializeWithValidDependencies()
    suiteResult.AddResult TestSendNotificationWithoutInitialize()
    suiteResult.AddResult TestSendNotificationWithInvalidParameters()
    suiteResult.AddResult TestSendNotificationConfigValuesUsed()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TINotificationServiceRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    
    ' Usar las constantes ya definidas para construir los nombres de archivo
    Dim templateDbName As String: templateDbName = "correos_test_template.accdb"
    Dim activeDbName As String: activeDbName = "correos_integration_test.accdb"
    
    ' Llamada al método correcto de modTestUtils
    modTestUtils.PrepareTestDatabase templateDbName, activeDbName
    
    Dim activePath As String: activePath = projectPath & "back\test_env\workspace\" & activeDbName
    Dim db As DAO.Database
    ' Abrir con la password que usan los tests de esta suite
    Set db = DBEngine.OpenDatabase(activePath, False, False, ";PWD=dpddpd")
    Call EnsureCorreosSchema(db)
    db.Close: Set db = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TINotificationService.SuiteSetup", Err.Description
End Sub

Private Sub SuiteTeardown()
    ' Limpieza centralizada usando CleanupTestDatabase
    modTestUtils.CleanupTestDatabase "correos_integration_test.accdb"
End Sub

Private Sub EnsureCorreosSchema(ByRef db As DAO.Database)
    Dim tdf As DAO.TableDef, exists As Boolean: exists = False
    For Each tdf In db.TableDefs
        If tdf.Name = "TbCorreosEnviados" Then exists = True: Exit For
    Next
    If Not exists Then
        db.Execute "CREATE TABLE TbCorreosEnviados (" & _
                   "Id AUTOINCREMENT PRIMARY KEY, " & _
                   "Destinatarios TEXT(255), " & _
                   "Asunto TEXT(255), " & _
                   "Cuerpo MEMO, " & _
                   "DestinatariosConCopia TEXT(255), " & _
                   "DestinatariosConCopiaOculta TEXT(255), " & _
                   "URLAdjunto TEXT(255), " & _
                   "FechaGrabacion DATETIME );", dbFailOnError
    End If
End Sub

' ============================================================================
' TESTS INDIVIDUALES (SE AÑADIRÁN EN LOS SIGUIENTES PROMPTS)
' ============================================================================

Private Function TestSendNotificationSuccessCallsRepositoryCorrectly() As CTestResult
    Set TestSendNotificationSuccessCallsRepositoryCorrectly = New CTestResult
    TestSendNotificationSuccessCallsRepositoryCorrectly.Initialize "SendNotification con éxito debe llamar al repositorio correctamente"
    
    Dim notificationService As INotificationService
    Dim db As DAO.Database
    
    On Error GoTo TestFail
    
    ' Crear configuración local apuntando a la BD de prueba de Correos
    Dim localConfig As IConfig
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "CORREOS_DB_PATH", modTestUtils.GetWorkspacePath() & "correos_integration_test.accdb"
     mockConfigImpl.SetSetting "CORREOS_PASSWORD", "dpddpd"
     mockConfigImpl.SetSetting "USUARIO_ACTUAL", "test.user@condor.com"
     mockConfigImpl.SetSetting "CORREO_ADMINISTRADOR", "admin@condor.com"
    Set localConfig = mockConfigImpl
    
    ' Crear el servicio real inyectando la configuración local
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(localConfig)
    Dim dbPassword As String: dbPassword = localConfig.GetCorreosPassword()
    Set db = DBEngine.OpenDatabase(modTestUtils.GetWorkspacePath() & "correos_integration_test.accdb", False, False, ";PWD=" & dbPassword)

    ' Act
    DBEngine.BeginTrans
    Dim success As Boolean
    success = notificationService.SendNotification("dest@empresa.com", "Asunto Test", "<html>Cuerpo Test</html>")
    
    ' Assert
    modAssert.AssertTrue success, "SendNotification debe retornar True en caso de éxito."
    ' Aquí podríamos añadir una aserción que verifique directamente en la BD que el correo se ha encolado.

    TestSendNotificationSuccessCallsRepositoryCorrectly.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationSuccessCallsRepositoryCorrectly.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set notificationService = Nothing
    Set db = Nothing
End Function

Private Function TestInitializeWithValidDependencies() As CTestResult
    Set TestInitializeWithValidDependencies = New CTestResult
    TestInitializeWithValidDependencies.Initialize "Initialize con dependencias válidas debe tener éxito"
    
    Dim notificationService As INotificationService
    On Error GoTo TestFail
    
    ' Arrange: El servicio obtiene la configuración del contexto centralizado
    ' Act: Intentar crear el servicio usando la factoría
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    ' Assert
    modAssert.AssertNotNull notificationService, "El servicio no debería ser nulo si las dependencias son válidas."
    
    TestInitializeWithValidDependencies.Pass
    GoTo Cleanup

TestFail:
    TestInitializeWithValidDependencies.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set notificationService = Nothing
End Function

Private Function TestSendNotificationWithoutInitialize() As CTestResult
    Set TestSendNotificationWithoutInitialize = New CTestResult
    TestSendNotificationWithoutInitialize.Initialize "SendNotification sin inicializar debe fallar devolviendo False"
    
    Dim notificationService As INotificationService
    On Error GoTo TestFail ' Si se produce un error de ejecución, el test falla.
    
    ' Arrange: Crear la instancia de la clase concreta pero SIN llamar a Initialize
    Dim notificationServiceImpl As New CNotificationService
    Set notificationService = notificationServiceImpl
    
    ' Act: Intentar usar el servicio no inicializado
    Dim success As Boolean
    success = notificationService.SendNotification("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Assert: El servicio debe fallar grácilmente devolviendo False, no con un error.
    modAssert.AssertFalse success, "SendNotification debe devolver False si el servicio no está inicializado."
    
    TestSendNotificationWithoutInitialize.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationWithoutInitialize.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set notificationService = Nothing
End Function

Private Function TestSendNotificationWithInvalidParameters() As CTestResult
    Set TestSendNotificationWithInvalidParameters = New CTestResult
    TestSendNotificationWithInvalidParameters.Initialize "SendNotification con parámetros inválidos debe devolver False"
    
    Dim notificationService As INotificationService
    On Error GoTo TestFail
    
    ' Arrange: El servicio obtiene la configuración del contexto centralizado
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    ' Act & Assert
    modAssert.AssertFalse notificationService.SendNotification("", "Asunto", "Cuerpo"), "Debe devolver False con destinatario vacío."
    modAssert.AssertFalse notificationService.SendNotification("test@test.com", "", "Cuerpo"), "Debe devolver False con asunto vacío."
    modAssert.AssertFalse notificationService.SendNotification("test@test.com", "Asunto", ""), "Debe devolver False con cuerpo vacío."
    
    TestSendNotificationWithInvalidParameters.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationWithInvalidParameters.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set notificationService = Nothing
End Function

Private Function TestSendNotificationConfigValuesUsed() As CTestResult
    Set TestSendNotificationConfigValuesUsed = New CTestResult
    TestSendNotificationConfigValuesUsed.Initialize "SendNotification debe usar los valores de configuración correctamente"
    
    Dim notificationService As INotificationService
    Dim db As DAO.Database
    
    On Error GoTo TestFail
    
    ' Arrange: Crear configuración local apuntando a la BD de prueba de Correos
    Dim localConfig As IConfig
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "CORREOS_DB_PATH", modTestUtils.GetWorkspacePath() & "correos_integration_test.accdb"
    mockConfigImpl.SetSetting "CORREOS_PASSWORD", "dpddpd"
    mockConfigImpl.SetSetting "USUARIO_ACTUAL", "test.user@condor.com"
    mockConfigImpl.SetSetting "CORREO_ADMINISTRADOR", "admin@condor.com"
    Set localConfig = mockConfigImpl
    
    ' Crear el servicio real inyectando la configuración local
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(localConfig)
    Dim dbPassword As String: dbPassword = localConfig.GetCorreosPassword()
    Set db = DBEngine.OpenDatabase(modTestUtils.GetWorkspacePath() & "correos_integration_test.accdb", False, False, ";PWD=" & dbPassword)

    ' Act
    DBEngine.BeginTrans
    Dim success As Boolean
    success = notificationService.SendNotification("dest@empresa.com", "Asunto Config Test", "<html>Test Config</html>")
    
    ' Assert
    modAssert.AssertTrue success, "SendNotification debe funcionar correctamente con config personalizado."
    
    TestSendNotificationConfigValuesUsed.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationConfigValuesUsed.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set notificationService = Nothing
    Set db = Nothing
End Function

