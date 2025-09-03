Attribute VB_Name = "TestCConfig"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CConfig
' Arquitectura: Pruebas Aisladas contra la implementación REAL de CConfig,
' utilizando LoadFromDictionary para inyectar configuración en memoria.
' ============================================================================

Public Function TestCConfigRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestCConfig - Pruebas Unitarias CConfig (Reconstruido)"
    
    suiteResult.AddResult TestGetDataPath_ReturnsCorrectValue()
    suiteResult.AddResult TestGetDatabasePassword_ReturnsCorrectValue()
    suiteResult.AddResult TestHasKey_ReturnsTrueForExistingKey()
    suiteResult.AddResult TestHasKey_ReturnsFalseForNonExistingKey()
    suiteResult.AddResult TestGetValue_ReturnsEmptyForNonExistingKey()
    
    Set TestCConfigRunAll = suiteResult
End Function

Private Function TestGetDataPath_ReturnsCorrectValue() As CTestResult
    Set TestGetDataPath_ReturnsCorrectValue = New CTestResult
    TestGetDataPath_ReturnsCorrectValue.Initialize "GetDataPath debe devolver el valor correcto cargado en memoria"
    
    Dim config As CConfig
    Dim testSettings As Scripting.Dictionary
    On Error GoTo TestFail

    ' Arrange
    Set config = New CConfig
    Set testSettings = New Scripting.Dictionary
    testSettings.CompareMode = TextCompare
    testSettings.Add "DATA_PATH", "C:\Ruta\De\Prueba.accdb"
    config.LoadFromDictionary testSettings

    ' Act
    Dim result As String
    result = config.GetDataPath()

    ' Assert
    modAssert.AssertEquals "C:\Ruta\De\Prueba.accdb", result, "GetDataPath no devolvió el valor inyectado."

    TestGetDataPath_ReturnsCorrectValue.Pass
    GoTo Cleanup

TestFail:
    TestGetDataPath_ReturnsCorrectValue.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
    Set testSettings = Nothing
End Function

Private Function TestGetDatabasePassword_ReturnsCorrectValue() As CTestResult
    Set TestGetDatabasePassword_ReturnsCorrectValue = New CTestResult
    TestGetDatabasePassword_ReturnsCorrectValue.Initialize "GetDatabasePassword debe devolver el valor correcto"
    
    Dim config As CConfig
    Dim testSettings As Scripting.Dictionary
    On Error GoTo TestFail

    ' Arrange
    Set config = New CConfig
    Set testSettings = New Scripting.Dictionary
    testSettings.CompareMode = TextCompare
    testSettings.Add "DATABASE_PASSWORD", "pass123"
    config.LoadFromDictionary testSettings

    ' Act
    Dim result As String
    result = config.GetDatabasePassword()

    ' Assert
    modAssert.AssertEquals "pass123", result, "GetDatabasePassword no devolvió el valor inyectado."

    TestGetDatabasePassword_ReturnsCorrectValue.Pass
    GoTo Cleanup

TestFail:
    TestGetDatabasePassword_ReturnsCorrectValue.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
    Set testSettings = Nothing
End Function

Private Function TestHasKey_ReturnsTrueForExistingKey() As CTestResult
    Set TestHasKey_ReturnsTrueForExistingKey = New CTestResult
    TestHasKey_ReturnsTrueForExistingKey.Initialize "HasKey debe devolver True para una clave existente"

    Dim config As CConfig
    Dim testSettings As Scripting.Dictionary
    On Error GoTo TestFail

    ' Arrange
    Set config = New CConfig
    Set testSettings = New Scripting.Dictionary
    testSettings.CompareMode = TextCompare
    testSettings.Add "EXISTING_KEY", "value"
    config.LoadFromDictionary testSettings

    ' Act
    Dim result As Boolean
    result = config.HasKey("EXISTING_KEY")

    ' Assert
    modAssert.AssertTrue result, "HasKey debería haber devuelto True."

    TestHasKey_ReturnsTrueForExistingKey.Pass
    GoTo Cleanup

TestFail:
    TestHasKey_ReturnsTrueForExistingKey.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
    Set testSettings = Nothing
End Function

Private Function TestHasKey_ReturnsFalseForNonExistingKey() As CTestResult
    Set TestHasKey_ReturnsFalseForNonExistingKey = New CTestResult
    TestHasKey_ReturnsFalseForNonExistingKey.Initialize "HasKey debe devolver False para una clave inexistente"
    
    Dim config As CConfig
    On Error GoTo TestFail

    ' Arrange
    Set config = New CConfig ' Sin cargar ningún diccionario

    ' Act
    Dim result As Boolean
    result = config.HasKey("NON_EXISTING_KEY")

    ' Assert
    modAssert.AssertFalse result, "HasKey debería haber devuelto False."

    TestHasKey_ReturnsFalseForNonExistingKey.Pass
    GoTo Cleanup

TestFail:
    TestHasKey_ReturnsFalseForNonExistingKey.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
End Function

Private Function TestGetValue_ReturnsEmptyForNonExistingKey() As CTestResult
    Set TestGetValue_ReturnsEmptyForNonExistingKey = New CTestResult
    TestGetValue_ReturnsEmptyForNonExistingKey.Initialize "GetValue debe devolver """" para una clave inexistente"

    Dim config As CConfig
    On Error GoTo TestFail

    ' Arrange
    Set config = New CConfig ' Sin cargar ningún diccionario

    ' Act
    Dim result As String
    result = config.GetValue("NON_EXISTING_KEY")

    ' Assert
    modAssert.AssertEquals "", result, "GetValue debería haber devuelto una cadena vacía."

TestGetValue_ReturnsEmptyForNonExistingKey.Pass
    GoTo Cleanup

TestFail:
    TestGetValue_ReturnsEmptyForNonExistingKey.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
End Function

' ============================================================================
' PRUEBAS UNITARIAS
' ============================================================================

