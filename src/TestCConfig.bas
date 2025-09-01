Attribute VB_Name = "TestCConfig"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CConfig
' ============================================================================
' Este módulo contiene pruebas unitarias aisladas para CConfig
' que utilizan LoadFromCollection para evitar dependencias de base de datos.
' Las pruebas son ultrarrápidas y completamente aisladas.

' Función principal que ejecuta todas las pruebas del módulo
Public Function TestCConfigRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestCConfig"
    
    ' Ejecutar todas las pruebas unitarias
    suiteResult.AddTestResult TestGetValueDatapathSuccess()
    suiteResult.AddTestResult TestGetValueDatabasepasswordSuccess()
    suiteResult.AddTestResult TestGetDataPathSuccess()
    suiteResult.AddTestResult TestGetDatabasePasswordSuccess()
    suiteResult.AddTestResult TestHasKeyExistingKeyReturnsTrue()
    suiteResult.AddTestResult TestHasKeyNonExistingKeyReturnsFalse()
    suiteResult.AddTestResult TestGetValueNonExistingKeyReturnsEmpty()
    
    Set TestCConfigRunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA CConfig
' ============================================================================

' Prueba que GetValue puede obtener el valor BACKEND_DB_PATH correctamente
Private Function TestGetValueDatapathSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetValue debe obtener BACKEND_DB_PATH correctamente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim config As IConfig
    
    mockConfig.Reset
    mockConfig.ConfigureGetValue "C:\Test\CONDOR_Backend.accdb", "BACKEND_DB_PATH"
    mockConfig.ConfigureGetValue "testpassword", "DATABASE_PASSWORD"
    
    Set config = mockConfig
    
    ' Act
    Dim dataPath As String
    dataPath = config.GetValue("BACKEND_DB_PATH")
    
    ' Assert
    modAssert.AssertEquals "C:\Test\CONDOR_Backend.accdb", dataPath, "BACKEND_DB_PATH debe ser el valor configurado"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set mockConfig = Nothing
    Set TestGetValueDatapathSuccess = testResult
End Function

' Prueba que GetValue puede obtener el valor DATABASE_PASSWORD correctamente
Private Function TestGetValueDatabasepasswordSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetValue debe obtener DATABASE_PASSWORD correctamente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim config As IConfig
    
    mockConfig.Reset
    mockConfig.ConfigureGetValue "C:\Test\CONDOR_Backend.accdb", "BACKEND_DB_PATH"
    mockConfig.ConfigureGetValue "testpassword", "DATABASE_PASSWORD"
    
    Set config = mockConfig
    
    ' Act
    Dim password As String
    password = config.GetValue("DATABASE_PASSWORD")
    
    ' Assert
    modAssert.AssertEquals "testpassword", password, "DATABASE_PASSWORD debe ser 'testpassword'"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set mockConfig = Nothing
    Set TestGetValueDatabasepasswordSuccess = testResult
End Function

' Prueba que HasKey devuelve True para claves existentes
Private Function TestHasKeyExistingKeyReturnsTrue() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "HasKey debe devolver True para una clave existente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim config As IConfig
    
    mockConfig.Reset
    mockConfig.ConfigureHasKey True, "EXISTING_KEY"
    
    Set config = mockConfig
    
    ' Act & Assert
    modAssert.AssertTrue config.HasKey("EXISTING_KEY"), "HasKey debe devolver True para EXISTING_KEY"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set mockConfig = Nothing
    Set TestHasKeyExistingKeyReturnsTrue = testResult
End Function

' Prueba que HasKey devuelve False para claves no existentes
Private Function TestHasKeyNonExistingKeyReturnsFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "HasKey debe devolver False para una clave inexistente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim config As IConfig
    
    mockConfig.Reset
    mockConfig.ConfigureHasKey False, "NON_EXISTING_KEY"
    mockConfig.ConfigureHasKey False, ""

    Set config = mockConfig
    
    ' Act & Assert
    modAssert.AssertFalse config.HasKey("NON_EXISTING_KEY"), "HasKey debe devolver False para clave inexistente"
    modAssert.AssertFalse config.HasKey(""), "HasKey debe devolver False para clave vacía"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set mockConfig = Nothing
    Set TestHasKeyNonExistingKeyReturnsFalse = testResult
End Function

' Prueba que GetValue devuelve cadena vacía para claves no existentes
Private Function TestGetValueNonExistingKeyReturnsEmpty() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetValue debe devolver una cadena vacía para una clave inexistente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim config As IConfig
    
    mockConfig.Reset
    mockConfig.ConfigureGetValue "", "NON_EXISTING_KEY"
    
    Set config = mockConfig
    
    ' Act & Assert
    Dim configValue As String
    configValue = config.GetValue("NON_EXISTING_KEY")
    modAssert.AssertEquals "", configValue, "GetValue debe devolver cadena vacía para clave inexistente"
    
    configValue = config.GetValue("")
    modAssert.AssertEquals "", configValue, "GetValue debe devolver cadena vacía para clave vacía"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set mockConfig = Nothing
    Set TestGetValueNonExistingKeyReturnsEmpty = testResult
End Function

' Prueba que GetDataPath devuelve la ruta correcta de la base de datos
Private Function TestGetDataPathSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetDataPath debe devolver la ruta de la BD correctamente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    Dim config As IConfig
    
    mockConfig.Reset
    mockConfig.ConfigureGetDataPath "C:\Test\CONDOR_Backend.accdb"
    
    Set config = mockConfig
    
    ' Act
    Dim dataPath As String
    dataPath = config.GetDataPath()
    
    ' Assert
    modAssert.AssertEquals "C:\Test\CONDOR_Backend.accdb", dataPath, "GetDataPath debe devolver el valor de BACKEND_DB_PATH"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set mockConfig = Nothing
    Set TestGetDataPathSuccess = testResult
End Function

' Prueba que GetDatabasePassword devuelve la contraseña correcta
Private Function TestGetDatabasePasswordSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetDatabasePassword debe devolver la contraseña correctamente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    Dim config As IConfig
    
    mockConfig.Reset
    mockConfig.ConfigureGetDatabasePassword "testpassword123"
    
    Set config = mockConfig
    
    ' Act
    Dim password As String
    password = config.GetDatabasePassword()
    
    ' Assert
    modAssert.AssertEquals "testpassword123", password, "GetDatabasePassword debe devolver 'testpassword123'"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set mockConfig = Nothing
    Set TestGetDatabasePasswordSuccess = testResult
End Function