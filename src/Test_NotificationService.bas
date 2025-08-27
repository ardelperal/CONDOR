Attribute VB_Name = "Test_NotificationService"
Option Compare Database
Option Explicit


#If DEV_MODE Then

' Test_NotificationService.bas
' Suite de pruebas unitarias PURAS para CNotificationService
' Refactorizado para usar patrón Repository con CMockNotificationRepository
' Pruebas 100% aisladas y rápidas

' Función principal para ejecutar todas las pruebas del módulo
Public Function Test_NotificationService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_NotificationService - Pruebas Unitarias CNotificationService"
    
    Debug.Print "=== Iniciando Test_NotificationService (REPOSITORY PATTERN) ==="
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result()
    suiteResult.AddTestResult Test_Initialize_WithValidDependencies_Result()
    suiteResult.AddTestResult Test_EnviarNotificacion_WithoutInitialize_Result()
    suiteResult.AddTestResult Test_EnviarNotificacion_WithInvalidParameters_Result()
    suiteResult.AddTestResult Test_EnviarNotificacion_ConfigValuesUsed_Result()
    
    Debug.Print "=== Test_NotificationService (REPOSITORY PATTERN) Completado ==="
    
    Set Test_NotificationService_RunAll = suiteResult
End Function

' Test: Verificar que EnviarNotificacion llama correctamente al repositorio
Private Function Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_Success_CallsRepositoryCorrectly"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Dim notificationServiceImpl As New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockNotificationRepository
    
    ' Configurar mock config con valores de prueba
    mockConfig.SetConfigValue "CORREOS_DB_PATH", "C:\test\correos_test.accdb"
    mockConfig.SetConfigValue "DATABASE_PASSWORD", "testpass123"
    mockConfig.SetConfigValue "USUARIO_ACTUAL", "testuser@empresa.com"
    mockConfig.SetConfigValue "CORREO_ADMINISTRADOR", "admin@empresa.com"
    
    ' Inicializar usando la implementación concreta
    notificationServiceImpl.Initialize mockConfig, mockLogger, mockRepository
    
    ' Asignar a la variable de interfaz para la prueba
    Set notificationService = notificationServiceImpl
    
    ' Act
    Dim Resultado As Boolean
    Resultado = notificationService.EnviarNotificacion("dest@empresa.com", "Asunto Test", "<html>Cuerpo Test</html>")
    
    ' Assert - Verificar que se llamó al repositorio correctamente
    Assert.IsTrue mockRepository.EncolarCorreo_WasCalled, "Debe llamar al método EncolarCorreo del repositorio"
    Assert.IsTrue Resultado, "EnviarNotificacion debe retornar True en caso de éxito"
    
    ' Verificar que se pasaron los parámetros correctos al repositorio
    Assert.AreEqual "Asunto Test", mockRepository.LastAsunto, "Debe pasar el asunto correcto"
    Assert.AreEqual "<html>Cuerpo Test</html>", mockRepository.LastCuerpo, "Debe pasar el cuerpo correcto"
    Assert.AreEqual "dest@empresa.com", mockRepository.LastDestinatarios, "Debe pasar los destinatarios correctos"
    Assert.AreEqual "testuser@empresa.com", mockRepository.LastCC, "Debe usar usuario actual como CC"
    Assert.AreEqual "admin@empresa.com", mockRepository.LastBCC, "Debe usar correo administrador como BCC"
    Assert.AreEqual "", mockRepository.LastAdjunto, "Debe pasar adjunto vacío por defecto"
    
    ' Verificar que se registró la operación en el logger
    Assert.IsTrue mockLogger.LogOperation_WasCalled, "Debe registrar la operación en el logger"
    Assert.AreEqual "NOTIFICATION", mockLogger.LastOperationType, "Debe registrar como operación NOTIFICATION"
    
    result.Passed = True
    result.message = "Prueba exitosa: EnviarNotificacion usa el repositorio correctamente"
    
    Set Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.message = "Error en Test_EnviarNotificacion_Success_CallsRepositoryCorrectly: " & Err.Description
    Set Test_EnviarNotificacion_Success_CallsRepositoryCorrectly_Result = result
End Function

' Test: Verificar inicialización correcta con todas las dependencias
Private Function Test_Initialize_WithValidDependencies_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_Initialize_WithValidDependencies"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Dim notificationServiceImpl As New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockNotificationRepository
    
    ' Inicializar usando la implementación concreta
    notificationServiceImpl.Initialize mockConfig, mockLogger, mockRepository
    
    ' Asignar a la variable de interfaz para la prueba
    Set notificationService = notificationServiceImpl
    
    ' Act - La inicialización ya se hizo arriba, aquí solo verificamos que no hay error
    
    ' Assert - Si no hay error, la inicialización fue exitosa
    result.Passed = True
    result.message = "Prueba exitosa: Initialize con todas las dependencias válidas"
    
    Set Test_Initialize_WithValidDependencies_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.message = "Error en Test_Initialize_WithValidDependencies: " & Err.Description
    Set Test_Initialize_WithValidDependencies_Result = result
End Function

' Test: Verificar que EnviarNotificacion falla sin inicializar
Private Function Test_EnviarNotificacion_WithoutInitialize_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_WithoutInitialize"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Dim notificationServiceImpl As New CNotificationService
    ' No llamamos Initialize intencionalmente
    Set notificationService = notificationServiceImpl
    
    ' Act
    Dim Resultado As Boolean
    Resultado = notificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Assert
    Assert.IsFalse Resultado, "EnviarNotificacion debe fallar sin inicializar"
    
    result.Passed = True
    result.message = "Prueba exitosa: EnviarNotificacion falla correctamente sin Initialize"
    
    Set Test_EnviarNotificacion_WithoutInitialize_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.message = "Error en Test_EnviarNotificacion_WithoutInitialize: " & Err.Description
    Set Test_EnviarNotificacion_WithoutInitialize_Result = result
End Function

' Test: Verificar comportamiento con parámetros inválidos
Private Function Test_EnviarNotificacion_WithInvalidParameters_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_WithInvalidParameters"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Dim notificationServiceImpl As New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockNotificationRepository
    
    ' Inicializar usando la implementación concreta
    notificationServiceImpl.Initialize mockConfig, mockLogger, mockRepository
    
    ' Asignar a la variable de interfaz para la prueba
    Set notificationService = notificationServiceImpl
    
    ' Act & Assert - Destinatarios vacío
    Dim resultado1 As Boolean
    resultado1 = notificationService.EnviarNotificacion("", "Asunto", "<html>Cuerpo</html>")
    Assert.IsFalse resultado1, "Debe fallar con destinatarios vacío"
    
    ' Act & Assert - Asunto vacío
    Dim resultado2 As Boolean
    resultado2 = notificationService.EnviarNotificacion("test@empresa.com", "", "<html>Cuerpo</html>")
    Assert.IsFalse resultado2, "Debe fallar con asunto vacío"
    
    ' Act & Assert - Cuerpo vacío
    Dim resultado3 As Boolean
    resultado3 = notificationService.EnviarNotificacion("test@empresa.com", "Asunto", "")
    Assert.IsFalse resultado3, "Debe fallar con cuerpo vacío"
    
    ' Verificar que el repositorio no fue llamado en ningún caso
    Assert.IsFalse mockRepository.EncolarCorreo_WasCalled, "No debe llamar al repositorio con parámetros inválidos"
    
    result.Passed = True
    result.message = "Prueba exitosa: EnviarNotificacion valida correctamente parámetros"
    
    Set Test_EnviarNotificacion_WithInvalidParameters_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.message = "Error en Test_EnviarNotificacion_WithInvalidParameters: " & Err.Description
    Set Test_EnviarNotificacion_WithInvalidParameters_Result = result
End Function

' Test: Verificar que se usan los valores correctos del config
Private Function Test_EnviarNotificacion_ConfigValuesUsed_Result() As CTestResult
    Dim result As New CTestResult
    result.testName = "Test_EnviarNotificacion_ConfigValuesUsed"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Dim notificationServiceImpl As New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockNotificationRepository
    
    ' Configurar valores específicos en el mock
    mockConfig.SetConfigValue "CORREOS_DB_PATH", "C:\test\correos_test.accdb"
    mockConfig.SetConfigValue "DATABASE_PASSWORD", "password123"
    mockConfig.SetConfigValue "USUARIO_ACTUAL", "usuario.test@empresa.com"
    mockConfig.SetConfigValue "CORREO_ADMINISTRADOR", "admin.test@empresa.com"
    
    ' Inicializar usando la implementación concreta
    notificationServiceImpl.Initialize mockConfig, mockLogger, mockRepository
    
    ' Asignar a la variable de interfaz para la prueba
    Set notificationService = notificationServiceImpl
    
    ' Act
    Dim Resultado As Boolean
    Resultado = notificationService.EnviarNotificacion("dest@empresa.com", "Asunto Config Test", "<html>Test Config</html>")
    
    ' Assert - Verificar que se obtuvieron los valores correctos del config
    Assert.IsTrue mockConfig.GetCorreosDBPath_WasCalled, "Debe obtener CORREOS_DB_PATH del config"
    Assert.IsTrue mockConfig.GetDatabasePassword_WasCalled, "Debe obtener DATABASE_PASSWORD del config"
    Assert.IsTrue mockConfig.GetUsuarioActual_WasCalled, "Debe obtener USUARIO_ACTUAL del config"
    Assert.IsTrue mockConfig.GetCorreoAdministrador_WasCalled, "Debe obtener CORREO_ADMINISTRADOR del config"
    
    ' Verificar que los valores del config se pasaron correctamente al repositorio
    Assert.AreEqual "usuario.test@empresa.com", mockRepository.LastCC, "Debe usar el usuario actual del config como CC"
    Assert.AreEqual "admin.test@empresa.com", mockRepository.LastBCC, "Debe usar el correo administrador del config como BCC"
    
    result.Passed = True
    result.message = "Prueba exitosa: EnviarNotificacion usa valores correctos del config inyectado"
    
    Set Test_EnviarNotificacion_ConfigValuesUsed_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.message = "Error en Test_EnviarNotificacion_ConfigValuesUsed: " & Err.Description
    Set Test_EnviarNotificacion_ConfigValuesUsed_Result = result
End Function

#End If








