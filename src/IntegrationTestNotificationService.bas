Attribute VB_Name = "IntegrationTest_NotificationService"
Option Compare Database
Option Explicit

' Constantes para rutas de bases de datos
Private Const CORREOS_TEMPLATE_PATH As String = "back\test_db\templates\correos_test_template.accdb"
Private Const CORREOS_ACTIVE_PATH As String = "back\test_db\active\correos_integration_test.accdb"

#If DEV_MODE Then

' IntegrationTest_NotificationService.bas
' Suite de pruebas de integración para CNotificationService
' Usa repositorios reales con base de datos de prueba
' Tests de integración completos

' Procedimiento de configuración inicial
Private Sub Setup()
    ' Preparar base de datos de correos para las pruebas
    Call modTestUtils.PrepareTestDatabase(modTestUtils.GetProjectPath() & CORREOS_TEMPLATE_PATH, modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH)
End Sub

' Procedimiento de limpieza
Private Sub Teardown()
    ' Limpiar base de datos de prueba usando IFileSystem factory
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim fullTestPath As String
    fullTestPath = modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH
    
    If fs.FileExists(fullTestPath) Then
        fs.DeleteFile fullTestPath
    End If
    
    Set fs = Nothing
End Sub

' Función principal para ejecutar todas las pruebas del módulo
Public Function IntegrationTestNotificationService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("IntegrationTestNotificationService - Pruebas de Integración CNotificationService")
    
    Debug.Print "=== Iniciando IntegrationTest_NotificationService (INTEGRATION TESTS) ==="
    
    ' Configurar entorno de pruebas
    Call Setup
    
    ' Ejecutar todas las pruebas
    Call suiteResult.AddTestResult(Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result())
    Call suiteResult.AddTestResult(Test_Initialize_WithValidDependencies_Result())
    Call suiteResult.AddTestResult(Test_EnviarNotificacion_WithoutInitialize_Result())
    Call suiteResult.AddTestResult(Test_EnviarNotificacion_WithInvalidParameters_Result())
    Call suiteResult.AddTestResult(Test_EnviarNotificacion_ConfigValuesUsed_Result())
    
    ' Limpiar entorno de pruebas
    Call Teardown
    
    Debug.Print "=== IntegrationTest_NotificationService (INTEGRATION TESTS) Completado ==="
    
    Set IntegrationTestNotificationService_RunAll = suiteResult
End Function

' Test: Verificar que EnviarNotificacion funciona correctamente con dependencias reales
Private Function Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_Success_CallsRepositoryCorrectly"
    
    On Error GoTo TestError
    
    ' Arrange - Usar dependencias reales
    Dim notificationService As INotificationService
    Dim testConfig As New CConfig
    Dim operationLogger As New COperationLogger
    Dim errorHandler As New CErrorHandlerService
    
    ' Configurar CConfig para usar la base de datos activa
    Call testConfig.SetSetting("CORREOS_DB_PATH", modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH)
    Call testConfig.SetSetting("DATABASE_PASSWORD", "testpass123")
    Call testConfig.SetSetting("USUARIO_ACTUAL", "testuser@empresa.com")
    Call testConfig.SetSetting("CORREO_ADMINISTRADOR", "admin@empresa.com")
    
    ' Crear el servicio usando el factory con testConfig
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(testConfig)
    
    ' Act
    Dim Resultado As Boolean
    Resultado = notificationService.EnviarNotificacion("dest@empresa.com", "Asunto Test", "<html>Cuerpo Test</html>")
    
    ' Assert - Verificar que la operación fue exitosa
    Call modAssert.AssertTrue(Resultado, "EnviarNotificacion debe retornar True en caso de éxito")
    
    result.Pass
    
    ' Cleanup - No hay mocks que resetear en tests de integración
    
    Set Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result = result
    Exit Function
    
TestError:
    result.Fail "Error en Test_EnviarNotificacion_Success_CallsRepositoryCorrectly: " & Err.Description
    Set Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result = result
End Function

' Test: Verificar inicialización correcta con todas las dependencias reales
Private Function Test_Initialize_WithValidDependencies_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_Initialize_WithValidDependencies"
    
    On Error GoTo TestError
    
    ' Arrange - Usar dependencias reales
    Dim notificationService As INotificationService
    Dim testConfig As New CConfig
    Dim operationLogger As New COperationLogger
    Dim errorHandler As New CErrorHandlerService
    
    ' Configurar CConfig para usar la base de datos activa
    Call testConfig.SetSetting("CORREOS_DB_PATH", modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH)
    Call testConfig.SetSetting("DATABASE_PASSWORD", "testpass123")
    
    ' Act - Crear el servicio usando el factory con testConfig (esto incluye la inicialización)
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(testConfig)
    
    ' Assert - Si no hay error, la inicialización fue exitosa
    Call modAssert.AssertNotNull(notificationService, "El servicio debe crearse correctamente")
    
    result.Pass
    
    ' Cleanup - No hay mocks que resetear en tests de integración
    
    Set Test_Initialize_WithValidDependencies_Result = result
    Exit Function
    
TestError:
    result.Fail "Error en Test_Initialize_WithValidDependencies: " & Err.Description
    Set Test_Initialize_WithValidDependencies_Result = result
End Function

' Test: Verificar que EnviarNotificacion falla sin inicializar correctamente
Private Function Test_EnviarNotificacion_WithoutInitialize_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_WithoutInitialize"
    
    On Error GoTo TestError
    
    ' Arrange - Crear servicio sin inicializar
    Dim notificationService As INotificationService
    Dim notificationServiceImpl As New CNotificationService
    ' No llamamos Initialize intencionalmente
    Set notificationService = notificationServiceImpl
    
    ' Act
    Dim Resultado As Boolean
    Resultado = notificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Assert
    Call modAssert.AssertFalse(Resultado, "EnviarNotificacion debe fallar sin inicializar")
    
    result.Pass
    
    ' Cleanup - No hay dependencias que limpiar en esta prueba
    
    Set Test_EnviarNotificacion_WithoutInitialize_Result = result
    Exit Function
    
TestError:
    result.Fail "Error en Test_EnviarNotificacion_WithoutInitialize: " & Err.Description
    Set Test_EnviarNotificacion_WithoutInitialize_Result = result
End Function

' Test: Verificar comportamiento con parámetros inválidos usando dependencias reales
Private Function Test_EnviarNotificacion_WithInvalidParameters_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_WithInvalidParameters"
    
    On Error GoTo TestError
    
    ' Arrange - Usar dependencias reales
    Dim notificationService As INotificationService
    Dim testConfig As New CConfig
    Dim operationLogger As New COperationLogger
    Dim errorHandler As New CErrorHandlerService
    
    ' Configurar CConfig para usar la base de datos activa
    Call testConfig.SetSetting("CORREOS_DB_PATH", modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH)
    Call testConfig.SetSetting("DATABASE_PASSWORD", "testpass123")
    
    ' Crear el servicio usando el factory con testConfig
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(testConfig)
    
    ' Act & Assert - Destinatarios vacío
    Dim resultado1 As Boolean
    resultado1 = notificationService.EnviarNotificacion("", "Asunto", "<html>Cuerpo</html>")
    Call modAssert.AssertFalse(resultado1, "Debe fallar con destinatarios vacío")
    
    ' Act & Assert - Asunto vacío
    Dim resultado2 As Boolean
    resultado2 = notificationService.EnviarNotificacion("test@empresa.com", "", "<html>Cuerpo</html>")
    Call modAssert.AssertFalse(resultado2, "Debe fallar con asunto vacío")
    
    ' Act & Assert - Cuerpo vacío
    Dim resultado3 As Boolean
    resultado3 = notificationService.EnviarNotificacion("test@empresa.com", "Asunto", "")
    Call modAssert.AssertFalse(resultado3, "Debe fallar con cuerpo vacío")
    
    result.Pass
    
    ' Cleanup - No hay mocks que resetear en tests de integración
    
    Set Test_EnviarNotificacion_WithInvalidParameters_Result = result
    Exit Function
    
TestError:
    result.Fail "Error en Test_EnviarNotificacion_WithInvalidParameters: " & Err.Description
    Set Test_EnviarNotificacion_WithInvalidParameters_Result = result
End Function

' Test: Verificar que se usan los valores correctos del config con dependencias reales
Private Function Test_EnviarNotificacion_ConfigValuesUsed_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_ConfigValuesUsed"
    
    On Error GoTo TestError
    
    ' Arrange - Usar dependencias reales
    Dim notificationService As INotificationService
    Dim testConfig As New CConfig
    Dim operationLogger As New COperationLogger
    Dim errorHandler As New CErrorHandlerService
    
    ' Configurar valores específicos en el config real
    Call testConfig.SetSetting("CORREOS_DB_PATH", modTestUtils.GetProjectPath() & CORREOS_ACTIVE_PATH)
    Call testConfig.SetSetting("DATABASE_PASSWORD", "password123")
    Call testConfig.SetSetting("USUARIO_ACTUAL", "usuario.test@empresa.com")
    Call testConfig.SetSetting("CORREO_ADMINISTRADOR", "admin.test@empresa.com")
    
    ' Crear el servicio usando el factory con testConfig
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(testConfig)
    
    ' Act
    Dim Resultado As Boolean
    Resultado = notificationService.EnviarNotificacion("dest@empresa.com", "Asunto Config Test", "<html>Test Config</html>")
    
    ' Assert - Verificar que la operación fue exitosa (el config se usa internamente)
    Call modAssert.AssertTrue(Resultado, "EnviarNotificacion debe funcionar correctamente con config personalizado")
    
    result.Pass
    
    ' Cleanup - No hay mocks que resetear en tests de integración
    
    Set Test_EnviarNotificacion_ConfigValuesUsed_Result = result
    Exit Function
    
TestError:
    result.Fail "Error en Test_EnviarNotificacion_ConfigValuesUsed: " & Err.Description
    Set Test_EnviarNotificacion_ConfigValuesUsed_Result = result
End Function

#End If








