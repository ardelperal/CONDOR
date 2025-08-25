Option Compare Database
Option Explicit

#If DEV_MODE Then

' Test_NotificationService.bas
' Suite de pruebas unitarias PURAS para CNotificationService
' Reconstruido para probar la implementación real con aislamiento completo
' Incluye CMockDatabase para verificar interacciones con la base de datos

' Función principal para ejecutar todas las pruebas del módulo
Public Function Test_NotificationService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.SuiteName = "Test_NotificationService"
    
    Debug.Print "=== Iniciando Test_NotificationService (PURO) ==="
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTest "Test_EnviarNotificacion_Success_CallsDatabaseCorrectly", Test_EnviarNotificacion_Success_CallsDatabaseCorrectly_Result()
    suiteResult.AddTest "Test_Initialize_WithValidDependencies", Test_Initialize_WithValidDependencies_Result()
    suiteResult.AddTest "Test_EnviarNotificacion_WithoutInitialize", Test_EnviarNotificacion_WithoutInitialize_Result()
    suiteResult.AddTest "Test_EnviarNotificacion_WithInvalidParameters", Test_EnviarNotificacion_WithInvalidParameters_Result()
    suiteResult.AddTest "Test_EnviarNotificacion_ConfigValuesUsed", Test_EnviarNotificacion_ConfigValuesUsed_Result()
    
    Debug.Print "=== Test_NotificationService (PURO) Completado ==="
    
    Set Test_NotificationService_RunAll = suiteResult
End Function

' Test: Verificar que EnviarNotificacion llama correctamente a la base de datos
Private Function Test_EnviarNotificacion_Success_CallsDatabaseCorrectly_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_Success_CallsDatabaseCorrectly"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Set notificationService = New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockDatabase As New CMockDatabase
    
    ' Configurar mock config con valores de prueba
    mockConfig.SetConfigValue "CORREOS_DB_PATH", "C:\test\correos_test.accdb"
    mockConfig.SetConfigValue "DATABASE_PASSWORD", "testpass123"
    mockConfig.SetConfigValue "USUARIO_ACTUAL", "testuser@empresa.com"
    mockConfig.SetConfigValue "CORREO_ADMINISTRADOR", "admin@empresa.com"
    
    ' Inyectar dependencias
    notificationService.Initialize mockConfig, mockLogger
    
    ' Configurar el mock database para simular éxito
    mockDatabase.SimulateSuccess = True
    
    ' Act
    Dim resultado As Boolean
    resultado = notificationService.EnviarNotificacion("dest@empresa.com", "Asunto Test", "<html>Cuerpo Test</html>")
    
    ' Assert - Verificar que se usó la base de datos correctamente
    Assert.IsTrue mockDatabase.WasCreateQueryDefCalled, "Debe llamar a CreateQueryDef para consulta parametrizada"
    Assert.IsTrue mockDatabase.WasExecuteCalled, "Debe llamar a Execute para ejecutar la consulta"
    
    ' Verificar que se creó la consulta con el nombre correcto
    Assert.AreEqual "qryInsertCorreo", mockDatabase.LastQueryDefName, "Debe crear QueryDef con nombre correcto"
    
    ' Verificar que la consulta SQL es parametrizada (no contiene valores directos)
    Dim sqlQuery As String
    sqlQuery = mockDatabase.LastQueryDefSQL
    Assert.IsTrue InStr(sqlQuery, "INSERT INTO") > 0, "Debe ser una consulta INSERT"
    Assert.IsTrue InStr(sqlQuery, "?") > 0, "Debe usar parámetros (?) en lugar de concatenación"
    Assert.IsFalse InStr(sqlQuery, "dest@empresa.com") > 0, "No debe contener valores literales (SQL injection prevention)"
    
    ' Verificar que se establecieron los parámetros correctamente
    Assert.AreEqual 10, mockDatabase.ParameterCount, "Debe tener 10 parámetros"
    Assert.AreEqual "dest@empresa.com", mockDatabase.GetParameterValue(0), "Parámetro 0 debe ser destinatarios"
    Assert.AreEqual "Asunto Test", mockDatabase.GetParameterValue(1), "Parámetro 1 debe ser asunto"
    Assert.AreEqual "<html>Cuerpo Test</html>", mockDatabase.GetParameterValue(2), "Parámetro 2 debe ser cuerpo"
    Assert.AreEqual "testuser@empresa.com", mockDatabase.GetParameterValue(3), "Parámetro 3 debe ser usuario actual"
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion usa consultas parametrizadas correctamente"
    
    Set Test_EnviarNotificacion_Success_CallsDatabaseCorrectly_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_Success_CallsDatabaseCorrectly: " & Err.Description
    Set Test_EnviarNotificacion_Success_CallsDatabaseCorrectly_Result = result
End Function

' Test: Verificar inicialización correcta con dependencias válidas
Private Function Test_Initialize_WithValidDependencies_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_Initialize_WithValidDependencies"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Set notificationService = New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    
    ' Act
    notificationService.Initialize mockConfig, mockLogger
    
    ' Assert - Si no hay error, la inicialización fue exitosa
    result.Passed = True
    result.Message = "Prueba exitosa: Initialize con dependencias válidas"
    
    Set Test_Initialize_WithValidDependencies_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_Initialize_WithValidDependencies: " & Err.Description
    Set Test_Initialize_WithValidDependencies_Result = result
End Function

' Test: Verificar que EnviarNotificacion falla sin inicializar
Private Function Test_EnviarNotificacion_WithoutInitialize_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_WithoutInitialize"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Set notificationService = New CNotificationService
    ' No llamamos Initialize intencionalmente
    
    ' Act
    Dim resultado As Boolean
    resultado = notificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Assert
    Assert.IsFalse resultado, "EnviarNotificacion debe fallar sin inicializar"
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion falla correctamente sin Initialize"
    
    Set Test_EnviarNotificacion_WithoutInitialize_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_WithoutInitialize: " & Err.Description
    Set Test_EnviarNotificacion_WithoutInitialize_Result = result
End Function

' Test: Verificar comportamiento con parámetros inválidos
Private Function Test_EnviarNotificacion_WithInvalidParameters_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_WithInvalidParameters"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Set notificationService = New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    
    notificationService.Initialize mockConfig, mockLogger
    
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
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion valida correctamente parámetros"
    
    Set Test_EnviarNotificacion_WithInvalidParameters_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_WithInvalidParameters: " & Err.Description
    Set Test_EnviarNotificacion_WithInvalidParameters_Result = result
End Function

' Test: Verificar que se usan los valores correctos del config
Private Function Test_EnviarNotificacion_ConfigValuesUsed_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_ConfigValuesUsed"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim notificationService As INotificationService
    Set notificationService = New CNotificationService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    
    ' Configurar valores específicos en el mock
    mockConfig.SetConfigValue "CORREOS_DB_PATH", "C:\test\correos_test.accdb"
    mockConfig.SetConfigValue "DATABASE_PASSWORD", "password123"
    mockConfig.SetConfigValue "USUARIO_ACTUAL", "usuario.test@empresa.com"
    mockConfig.SetConfigValue "CORREO_ADMINISTRADOR", "admin.test@empresa.com"
    
    notificationService.Initialize mockConfig, mockLogger
    
    ' Act
    Dim resultado As Boolean
    resultado = notificationService.EnviarNotificacion("dest@empresa.com", "Asunto Config Test", "<html>Test Config</html>")
    
    ' Assert - Verificar que se obtuvieron los valores correctos
    Assert.IsTrue mockConfig.GetCorreosDBPath_WasCalled, "Debe obtener CORREOS_DB_PATH del config"
    Assert.IsTrue mockConfig.GetDatabasePassword_WasCalled, "Debe obtener DATABASE_PASSWORD del config"
    Assert.IsTrue mockConfig.GetUsuarioActual_WasCalled, "Debe obtener USUARIO_ACTUAL del config"
    Assert.IsTrue mockConfig.GetCorreoAdministrador_WasCalled, "Debe obtener CORREO_ADMINISTRADOR del config"
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion usa valores correctos del config inyectado"
    
    Set Test_EnviarNotificacion_ConfigValuesUsed_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_ConfigValuesUsed: " & Err.Description
    Set Test_EnviarNotificacion_ConfigValuesUsed_Result = result
End Function

#End If






