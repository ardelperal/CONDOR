Attribute VB_Name = "TINotificationService"
Option Compare Database
Option Explicit


' Constantes para rutas de bases de datos
Private Const CORREOS_TEMPLATE_PATH As String = "back\test_db\templates\correos_test_template.accdb"
Private Const CORREOS_ACTIVE_PATH As String = "back\test_db\active\correos_integration_test.accdb"

' Suite de pruebas de integración para CNotificationService

Private Sub Setup()
    Call modTestUtils.PrepareTestDatabase(modTestUtils.GetProjectPath() & CORREOS_TEMPLATE_PATH, modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH)
End Sub

Private Sub Teardown()
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim fullTestPath As String
    fullTestPath = modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH
    If fs.FileExists(fullTestPath) Then
        fs.DeleteFile fullTestPath
    End If
    Set fs = Nothing
End Sub

Public Function TINotificationServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TINotificationService")
    
    Call suiteResult.AddResult(TestSendNotificationSuccessCallsRepositoryCorrectly())
    Call suiteResult.AddResult(TestInitializeWithValidDependencies())
    Call suiteResult.AddResult(TestSendNotificationWithoutInitialize())
    Call suiteResult.AddResult(TestSendNotificationWithInvalidParameters())
    Call suiteResult.AddResult(TestSendNotificationConfigValuesUsed())
    
    Set TINotificationServiceRunAll = suiteResult
End Function

Private Function TestSendNotificationSuccessCallsRepositoryCorrectly() As CTestResult
    Dim result As New CTestResult
    result.Initialize "SendNotification con éxito debe llamar al repositorio correctamente"
    On Error GoTo TestError
    
    Call Setup
    
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    Dim success As Boolean
    success = notificationService.SendNotification("dest@empresa.com", "Asunto Test", "<html>Cuerpo Test</html>")
    
    Call modAssert.AssertTrue(success, "SendNotification debe retornar True en caso de éxito")
    
    result.Pass
    GoTo Cleanup
    
TestError:
    result.Fail "Error en TestSendNotificationSuccessCallsRepositoryCorrectly: " & Err.Description
Cleanup:
    Call Teardown
    Set TestSendNotificationSuccessCallsRepositoryCorrectly = result
End Function

Private Function TestInitializeWithValidDependencies() As CTestResult
    Dim result As New CTestResult
    result.Initialize "Initialize con dependencias válidas debe tener éxito"
    On Error GoTo TestError
    
    Call Setup
    
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    Call modAssert.AssertNotNull(notificationService, "El servicio debe crearse correctamente")
    
    result.Pass
    GoTo Cleanup

TestError:
    result.Fail "Error en TestInitializeWithValidDependencies: " & Err.Description
Cleanup:
    Call Teardown
    Set TestInitializeWithValidDependencies = result
End Function

Private Function TestSendNotificationWithoutInitialize() As CTestResult
    Dim result As New CTestResult
    result.Initialize "SendNotification sin inicializar debe fallar"
    On Error GoTo TestError
    
    Call Setup
    
    Dim notificationService As INotificationService
    Dim notificationServiceImpl As New CNotificationService
    Set notificationService = notificationServiceImpl
    
    Dim success As Boolean
    success = notificationService.SendNotification("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    Call modAssert.AssertFalse(success, "SendNotification debe fallar sin inicializar")
    
    result.Pass
    GoTo Cleanup

TestError:
    result.Fail "Error en TestSendNotificationWithoutInitialize: " & Err.Description
Cleanup:
    Call Teardown
    Set TestSendNotificationWithoutInitialize = result
End Function

Private Function TestSendNotificationWithInvalidParameters() As CTestResult
    Dim result As New CTestResult
    result.Initialize "SendNotification con parámetros inválidos debe fallar"
    On Error GoTo TestError
    
    Call Setup
    
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    Call modAssert.AssertFalse(notificationService.SendNotification("", "Asunto", "<html>Cuerpo</html>"), "Debe fallar con destinatarios vacío")
    Call modAssert.AssertFalse(notificationService.SendNotification("test@empresa.com", "", "<html>Cuerpo</html>"), "Debe fallar con asunto vacío")
    Call modAssert.AssertFalse(notificationService.SendNotification("test@empresa.com", "Asunto", ""), "Debe fallar con cuerpo vacío")
    
    result.Pass
    GoTo Cleanup

TestError:
    result.Fail "Error en TestSendNotificationWithInvalidParameters: " & Err.Description
Cleanup:
    Call Teardown
    Set TestSendNotificationWithInvalidParameters = result
End Function

Private Function TestSendNotificationConfigValuesUsed() As CTestResult
    Dim result As New CTestResult
    result.Initialize "SendNotification debe usar los valores de configuración correctamente"
    On Error GoTo TestError
    
    Call Setup
    
    Dim notificationService As INotificationService
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    Dim success As Boolean
    success = notificationService.SendNotification("dest@empresa.com", "Asunto Config Test", "<html>Test Config</html>")
    
    Call modAssert.AssertTrue(success, "SendNotification debe funcionar correctamente con config personalizado")
    
    result.Pass
    GoTo Cleanup

TestError:
    result.Fail "Error en TestSendNotificationConfigValuesUsed: " & Err.Description
Cleanup:
    Call Teardown
    Set TestSendNotificationConfigValuesUsed = result
End Function

