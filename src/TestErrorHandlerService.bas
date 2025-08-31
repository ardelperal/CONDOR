Attribute VB_Name = "TestErrorHandlerService"
Option Compare Database
Option Explicit

' =====================================================
' TestErrorHandlerService.bas
' Módulo de pruebas para CErrorHandlerService
' Implementa pruebas unitarias aisladas usando mocks
' =====================================================

' Función principal para ejecutar todas las pruebas del ErrorHandlerService
Public Function TestErrorHandlerServiceRunAll() As CTestSuiteResult
    On Error GoTo TestFail
    
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestErrorHandlerService"
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult (TestLogErrorWritesToFileSuccess())
    suiteResult.AddTestResult (TestLogErrorIsCriticalFlaggedCorrectly())
    suiteResult.AddTestResult (TestLogInfoWritesCorrectly())
    suiteResult.AddTestResult (TestLogWarningWritesCorrectly())
    suiteResult.AddTestResult (TestInitializeWithValidConfigSuccess())
    
    Set TestErrorHandlerServiceRunAll = suiteResult
    Exit Function
    
TestFail:
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "TestErrorHandlerService.TestErrorHandlerServiceRunAll", False
    Set TestErrorHandlerServiceRunAll = Nothing
End Function

' Prueba principal: LogError escribe correctamente usando mock
Public Function TestLogErrorWritesToFileSuccess() As CTestResult
    On Error GoTo TestFail
    
    Dim testResult As New CTestResult
    testResult.Initialize "LogError debe escribir en el fichero de log correctamente"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFileSystem As New CMockFileSystem
    Dim mockTextFile As CMockTextFile
    Dim writtenContent As String
    
    ' Arrange: Configurar el mock y el servicio
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\temp\test.log"
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    ' Inicializar el servicio con mocks
    errorHandlerService.Initialize mockConfig, mockFileSystem
    
    ' Act: Llamar al método LogError
    errorHandlerService.LogError 1001, "Error de prueba", "TestModule.TestFunction"
    
    ' Assert: Verificar que el mock capturó el contenido esperado
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "1001") > 0, "El número de error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "Error de prueba") > 0, "La descripción del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "TestModule.TestFunction") > 0, "La fuente del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "ERROR") > 0, "El nivel ERROR debe estar en el log")
    
    testResult.Pass
    
    Set TestLogErrorWritesToFileSuccess = testResult
    Exit Function
    
TestFail:
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "TestErrorHandlerService.TestLogErrorWritesToFileSuccess", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function

'Prueba: LogError con isCritical=True se marca correctamente en el JSON
Public Function TestLogErrorIsCriticalFlaggedCorrectly() As CTestResult
    On Error GoTo TestFail
    
    Dim testResult As New CTestResult
    testResult.Initialize "LogError con isCritical=True debe marcarse correctamente"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFileSystem As New CMockFileSystem
    Dim mockTextFile As CMockTextFile
    Dim writtenContent As String
    
    ' Arrange: Configurar el mock y el servicio
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\temp\test.log"
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    ' Inicializar el servicio con mocks
    errorHandlerService.Initialize mockConfig, mockFileSystem
    
    ' Act: Llamar al método LogError con isCritical = True
    errorHandlerService.LogError 2001, "Error crítico de prueba", "TestModule.TestFunction", True
    
    ' Assert: Verificar que el JSON contiene "isCritical":true
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(InStr(writtenContent, """isCritical"" : true") > 0, "El JSON debe contener 'isCritical':true")
    Call modAssert.IsTrue(InStr(writtenContent, "Error crítico de prueba") > 0, "La descripción del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "ERROR") > 0, "El nivel ERROR debe estar en el log")
    
    testResult.Pass
    
    Set TestLogErrorIsCriticalFlaggedCorrectly = testResult
    Exit Function
    
TestFail:
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "TestErrorHandlerService.TestLogErrorIsCriticalFlaggedCorrectly", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function

'Prueba: LogInfo escribe correctamente usando mock
Public Function TestLogInfoWritesCorrectly() As CTestResult
    On Error GoTo TestFail
    
    Dim testResult As New CTestResult
    testResult.Initialize "LogInfo debe escribir en el fichero de log correctamente"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFileSystem As New CMockFileSystem
    Dim mockTextFile As CMockTextFile
    Dim writtenContent As String
    
    ' Arrange
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\temp\test.log"
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    errorHandlerService.Initialize mockConfig, mockFileSystem
    
    ' Act
    errorHandlerService.LogInfo "Información de prueba", "TestModule.TestFunction"
    
    ' Assert
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "Información de prueba") > 0, "El mensaje de info debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "INFO") > 0, "El nivel INFO debe estar en el log")
    
    testResult.Pass
    
    Set TestLogInfoWritesCorrectly = testResult
    Exit Function
    
TestFail:
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "TestErrorHandlerService.TestLogInfoWritesCorrectly", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function

'Prueba: LogWarning escribe correctamente usando mock
Public Function TestLogWarningWritesCorrectly() As CTestResult
    On Error GoTo TestFail
    
    Dim testResult As New CTestResult
    testResult.Initialize "LogWarning debe escribir en el fichero de log correctamente"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFileSystem As New CMockFileSystem
    Dim mockTextFile As CMockTextFile
    Dim writtenContent As String
    
    ' Arrange
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\temp\test.log"
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    errorHandlerService.Initialize mockConfig, mockFileSystem
    
    ' Act
    errorHandlerService.LogWarning "Advertencia de prueba", "TestModule.TestFunction"
    
    ' Assert
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "Advertencia de prueba") > 0, "El mensaje de warning debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "WARNING") > 0, "El nivel WARNING debe estar en el log")
    
    testResult.Pass
    
    Set TestLogWarningWritesCorrectly = testResult
    Exit Function
    
TestFail:
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "TestErrorHandlerService.TestLogWarningWritesCorrectly", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function

'Prueba: Initialize funciona correctamente con configuración válida y mock
Public Function TestInitializeWithValidConfigSuccess() As CTestResult
    On Error GoTo TestFail
    
    Dim testResult As New CTestResult
    testResult.Initialize "Initialize debe completarse exitosamente con configuración válida"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFileSystem As New CMockFileSystem
    
    ' Arrange: Configurar el mock config
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\temp\test.log"
    
    ' Act
    errorHandlerService.Initialize mockConfig, mockFileSystem
    
    ' Assert: Si no hay error, la inicialización fue exitosa
    testResult.Pass
    
    Set TestInitializeWithValidConfigSuccess = testResult
    Exit Function
    
TestFail:
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "TestErrorHandlerService.TestInitializeWithValidConfigSuccess", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function