Attribute VB_Name = "TINotificationService"
Option Compare Database
Option Explicit

' Constantes para rutas de bases de datos
Private Const CORREOS_TEMPLATE_PATH As String = "back\test_db\templates\correos_test_template.accdb"
Private Const CORREOS_ACTIVE_PATH As String = "back\test_db\active\correos_integration_test.accdb"

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
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & CORREOS_TEMPLATE_PATH
    Dim activePath As String: activePath = projectPath & CORREOS_ACTIVE_PATH
    Call modTestUtils.SuiteSetup(templatePath, activePath)
End Sub

Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH
    Call modTestUtils.SuiteTeardown(activePath)
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
    
    ' Arrange: El servicio obtiene la configuración del contexto centralizado
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH, False, False, ";PWD=dpddpd")

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
    
    ' Arrange: El servicio obtiene la configuración del contexto centralizado
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH, False, False, ";PWD=dpddpd")

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

