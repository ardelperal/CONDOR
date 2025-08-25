Attribute VB_Name = "Test_CConfig"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' MÃ"DULO DE PRUEBAS UNITARIAS PARA CConfig
' ============================================================================
' Este mÃ³dulo contiene pruebas unitarias aisladas para CConfig
' que utilizan LoadFromCollection para evitar dependencias de base de datos.
' Las pruebas son ultrarrápidas y completamente aisladas.

' FunciÃ³n principal que ejecuta todas las pruebas del mÃ³dulo
Public Function Test_CConfig_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CConfig"
    
    ' Ejecutar todas las pruebas de integraciÃ³n
    suiteResult.AddTestResult Test_GetValue_DATAPATH_Success()
    suiteResult.AddTestResult Test_GetValue_DATABASEPASSWORD_Success()
    suiteResult.AddTestResult Test_GetDataPath_Success()
    suiteResult.AddTestResult Test_GetDatabasePassword_Success()
    suiteResult.AddTestResult Test_HasKey_ExistingKey_ReturnsTrue()
    suiteResult.AddTestResult Test_HasKey_NonExistingKey_ReturnsFalse()
    suiteResult.AddTestResult Test_GetValue_NonExistingKey_ReturnsEmpty()
    
    Set Test_CConfig_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÃ“N PARA CConfig
' ============================================================================

' Prueba que CConfig puede obtener el valor DATAPATH correctamente
Private Function Test_GetValue_DATAPATH_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetValue_DATAPATH_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    Dim settings As New Collection
    settings.Add "C:\Test\CONDOR_Backend.accdb", "DATAPATH"
    settings.Add "testpassword", "DATABASEPASSWORD"
    config.LoadFromCollection settings
    
    ' Act
    Dim dataPath As String
    dataPath = config.GetValue("DATAPATH")
    
    ' Assert
    modAssert.AssertEquals "C:\Test\CONDOR_Backend.accdb", dataPath, "DATAPATH debe ser el valor configurado"
    modAssert.AssertTrue InStr(dataPath, ".accdb") > 0, "DATAPATH debe contener .accdb"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_GetValue_DATAPATH_Success = testResult
End Function

' Prueba que CConfig puede obtener el valor DATABASEPASSWORD correctamente
Private Function Test_GetValue_DATABASEPASSWORD_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetValue_DATABASEPASSWORD_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    Dim settings As New Collection
    settings.Add "C:\Test\CONDOR_Backend.accdb", "DATAPATH"
    settings.Add "testpassword", "DATABASEPASSWORD"
    config.LoadFromCollection settings
    
    ' Act
    Dim password As String
    password = config.GetValue("DATABASEPASSWORD")
    
    ' Assert
    modAssert.AssertEquals "testpassword", password, "DATABASEPASSWORD debe ser 'testpassword'"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_GetValue_DATABASEPASSWORD_Success = testResult
End Function

' Prueba que HasKey devuelve True para claves existentes
Private Function Test_HasKey_ExistingKey_ReturnsTrue() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_HasKey_ExistingKey_ReturnsTrue"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    Dim settings As New Collection
    settings.Add "C:\Test\CONDOR_Backend.accdb", "DATAPATH"
    settings.Add "testpassword", "DATABASEPASSWORD"
    config.LoadFromCollection settings
    
    ' Act & Assert
    modAssert.AssertTrue config.HasKey("DATAPATH"), "HasKey debe devolver True para DATAPATH"
    modAssert.AssertTrue config.HasKey("DATABASEPASSWORD"), "HasKey debe devolver True para DATABASEPASSWORD"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_HasKey_ExistingKey_ReturnsTrue = testResult
End Function

' Prueba que HasKey devuelve False para claves no existentes
Private Function Test_HasKey_NonExistingKey_ReturnsFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_HasKey_NonExistingKey_ReturnsFalse"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    Dim settings As New Collection
    settings.Add "C:\Test\CONDOR_Backend.accdb", "DATAPATH"
    settings.Add "testpassword", "DATABASEPASSWORD"
    config.LoadFromCollection settings
    
    ' Act & Assert
    modAssert.AssertFalse config.HasKey("CLAVE_INEXISTENTE"), "HasKey debe devolver False para clave inexistente"
    modAssert.AssertFalse config.HasKey(""), "HasKey debe devolver False para clave vacÃ­a"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_HasKey_NonExistingKey_ReturnsFalse = testResult
End Function

' Prueba que GetValue devuelve cadena vacÃ­a para claves no existentes
Private Function Test_GetValue_NonExistingKey_ReturnsEmpty() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetValue_NonExistingKey_ReturnsEmpty"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    Dim settings As New Collection
    settings.Add "C:\Test\CONDOR_Backend.accdb", "DATAPATH"
    settings.Add "testpassword", "DATABASEPASSWORD"
    config.LoadFromCollection settings
    
    ' Act & Assert
    Dim value As String
    value = config.GetValue("CLAVE_INEXISTENTE")
    modAssert.AssertEquals "", value, "GetValue debe devolver cadena vacÃ­a para clave inexistente"
    
    value = config.GetValue("")
    modAssert.AssertEquals "", value, "GetValue debe devolver cadena vacÃ­a para clave vacÃ­a"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_GetValue_NonExistingKey_ReturnsEmpty = testResult
End Function

' Prueba que GetDataPath devuelve la ruta correcta de la base de datos
Private Function Test_GetDataPath_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetDataPath_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    Dim settings As New Collection
    settings.Add "C:\Test\CONDOR_Backend.accdb", "BACKEND_DB_PATH"
    settings.Add "testpassword", "DATABASE_PASSWORD"
    config.LoadFromCollection settings
    
    ' Act
    Dim dataPath As String
    dataPath = config.GetDataPath()
    
    ' Assert
    modAssert.AssertEquals "C:\Test\CONDOR_Backend.accdb", dataPath, "GetDataPath debe devolver el valor de BACKEND_DB_PATH"
    modAssert.AssertTrue InStr(dataPath, ".accdb") > 0, "GetDataPath debe contener .accdb"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_GetDataPath_Success = testResult
End Function

' Prueba que GetDatabasePassword devuelve la contraseña correcta
Private Function Test_GetDatabasePassword_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetDatabasePassword_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    Dim settings As New Collection
    settings.Add "C:\Test\CONDOR_Backend.accdb", "BACKEND_DB_PATH"
    settings.Add "testpassword123", "DATABASE_PASSWORD"
    config.LoadFromCollection settings
    
    ' Act
    Dim password As String
    password = config.GetDatabasePassword()
    
    ' Assert
    modAssert.AssertEquals "testpassword123", password, "GetDatabasePassword debe devolver 'testpassword123'"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_GetDatabasePassword_Success = testResult
End Function

#End If







