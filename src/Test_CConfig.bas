Attribute VB_Name = "Test_CConfig"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' MÓDULO DE PRUEBAS DE INTEGRACIÓN PARA CConfig
' ============================================================================
' Este módulo contiene pruebas de integración para CConfig
' que validan la nueva implementación autónoma sin dependencias.
' CConfig ahora se conecta directamente a la base de datos.

' Función principal que ejecuta todas las pruebas del módulo
Public Function Test_CConfig_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CConfig"
    
    ' Ejecutar todas las pruebas de integración
    suiteResult.AddTestResult Test_GetValue_DATAPATH_Success()
    suiteResult.AddTestResult Test_GetValue_DATABASEPASSWORD_Success()
    suiteResult.AddTestResult Test_HasKey_ExistingKey_ReturnsTrue()
    suiteResult.AddTestResult Test_HasKey_NonExistingKey_ReturnsFalse()
    suiteResult.AddTestResult Test_GetValue_NonExistingKey_ReturnsEmpty()
    
    Set Test_CConfig_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN PARA CConfig
' ============================================================================

' Prueba que CConfig puede obtener el valor DATAPATH correctamente
Private Function Test_GetValue_DATAPATH_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetValue_DATAPATH_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    
    ' Act
    Dim dataPath As String
    dataPath = config.GetValue("DATAPATH")
    
    ' Assert
    modAssert.AssertFalse dataPath = "", "DATAPATH no debe estar vacío"
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
    
    ' Act
    Dim password As String
    password = config.GetValue("DATABASEPASSWORD")
    
    ' Assert
    modAssert.AssertEquals "dpddpd", password, "DATABASEPASSWORD debe ser 'dpddpd'"
    
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
    
    ' Act & Assert
    modAssert.AssertFalse config.HasKey("CLAVE_INEXISTENTE"), "HasKey debe devolver False para clave inexistente"
    modAssert.AssertFalse config.HasKey(""), "HasKey debe devolver False para clave vacía"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_HasKey_NonExistingKey_ReturnsFalse = testResult
End Function

' Prueba que GetValue devuelve cadena vacía para claves no existentes
Private Function Test_GetValue_NonExistingKey_ReturnsEmpty() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetValue_NonExistingKey_ReturnsEmpty"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim config As New CConfig
    
    ' Act & Assert
    Dim value As String
    value = config.GetValue("CLAVE_INEXISTENTE")
    modAssert.AssertEquals "", value, "GetValue debe devolver cadena vacía para clave inexistente"
    
    value = config.GetValue("")
    modAssert.AssertEquals "", value, "GetValue debe devolver cadena vacía para clave vacía"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set config = Nothing
    Set Test_GetValue_NonExistingKey_ReturnsEmpty = testResult
End Function

#End If







