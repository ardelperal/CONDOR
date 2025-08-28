Attribute VB_Name = "Test_ErrorHandlerService"
Option Compare Database
Option Explicit



#If DEV_MODE Then

' =====================================================
' Test_ErrorHandlerService.bas
' Módulo de pruebas para CErrorHandlerService
' Implementa pruebas unitarias aisladas usando mocks
' =====================================================

' Función principal para ejecutar todas las pruebas del ErrorHandlerService
Public Function TestErrorHandlerServiceRunAll() As CTestSuiteResult
    On Error GoTo TestFail
    
    Dim ErrorHandler As IErrorHandlerService
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_ErrorHandlerService"
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult (Test_LogError_WritesToFile_Success())
    suiteResult.AddTestResult (Test_LogError_IsCritical_FlaggedCorrectly())
    suiteResult.AddTestResult (Test_LogInfo_WritesCorrectly())
    suiteResult.AddTestResult (Test_LogWarning_WritesCorrectly())
    suiteResult.AddTestResult (Test_Initialize_WithValidConfig_Success())
    
    Set TestErrorHandlerServiceRunAll = suiteResult
    Exit Function
    
TestFail:
    Dim mockConfig As New CMockConfig
    Dim mockFileSystem As New CMockFileSystem
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "Test_ErrorHandlerService.Test_ErrorHandlerService_RunAll", False
    Set Test_ErrorHandlerService_RunAll = Nothing
End Function

' Prueba principal: LogError escribe correctamente usando mock
Public Function TestLogErrorWritesToFileSuccess() As CTestResult
    On Error GoTo TestFail
    
    Dim ErrorHandler As IErrorHandlerService
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogError_WritesToFile_Success"
    
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
    errorHandlerService.LogError 1001, "Error de prueba", "Test_Module.Test_Function"
    
    ' Assert: Verificar que el mock capturó el contenido esperado
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "1001") > 0, "El número de error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "Error de prueba") > 0, "La descripción del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "Test_Module.Test_Function") > 0, "La fuente del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "ERROR") > 0, "El nivel ERROR debe estar en el log")
    
    testResult.Pass "LogError escribió correctamente usando mock"
    
    Set TestLogErrorWritesToFileSuccess = testResult
    Exit Function
    
TestFail:
    Dim mockConfigErr As New CMockConfig
    Dim mockFileSystemErr As New CMockFileSystem
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogError_WritesToFile_Success", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function

' Prueba: LogError con isCritical=True se marca correctamente en el JSON
Public Function Test_LogError_IsCritical_FlaggedCorrectly() As CTestResult
    On Error GoTo TestFail
    
    Dim ErrorHandler As IErrorHandlerService
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogError_IsCritical_FlaggedCorrectly"
    
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
    errorHandlerService.LogError 2001, "Error crítico de prueba", "Test_Module.Test_Function", True
    
    ' Assert: Verificar que el JSON contiene "isCritical":true
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(InStr(writtenContent, """isCritical"" : true") > 0, "El JSON debe contener 'isCritical':true")
    Call modAssert.IsTrue(InStr(writtenContent, "Error crítico de prueba") > 0, "La descripción del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "ERROR") > 0, "El nivel ERROR debe estar en el log")
    
    testResult.Pass "LogError con isCritical=True se marcó correctamente en el JSON"
    
    Set Test_LogError_IsCritical_FlaggedCorrectly = testResult
    Exit Function
    
TestFail:
    Dim mockConfigErr2 As New CMockConfig
    Dim mockFileSystemErr2 As New CMockFileSystem
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogError_IsCritical_FlaggedCorrectly", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function

' Prueba: LogInfo escribe correctamente usando mock
Public Function Test_LogInfo_WritesCorrectly() As CTestResult
    On Error GoTo TestFail
    
    Dim ErrorHandler As IErrorHandlerService
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogInfo_WritesCorrectly"
    
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
    errorHandlerService.LogInfo "Información de prueba", "Test_Module.Test_Function"
    
    ' Assert
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "Información de prueba") > 0, "El mensaje de info debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "INFO") > 0, "El nivel INFO debe estar en el log")
    
    testResult.Pass "LogInfo escribió correctamente al archivo de log"
    
    Set Test_LogInfo_WritesCorrectly = testResult
    Exit Function
    
TestFail:
    Dim mockConfigErr3 As New CMockConfig
    Dim mockFileSystemErr3 As New CMockFileSystem
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogWarning_WritesCorrectly", False
    testResult.Fail "Error durante la prueba: " & Err.Description
End Function

' Prueba: LogWarning escribe correctamente usando mock
Public Function Test_LogWarning_WritesCorrectly() As CTestResult
    On Error GoTo TestFail
    
    Dim ErrorHandler As IErrorHandlerService
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogWarning_WritesCorrectly"
    
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
    errorHandlerService.LogWarning "Advertencia de prueba", "Test_Module.Test_Function"
    
    ' Assert
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "Advertencia de prueba") > 0, "El mensaje de warning debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "WARNING") > 0, "El nivel WARNING debe estar en el log")
    
    testResult.Pass "LogWarning escribió correctamente usando mock"
    
    Set Test_LogWarning_WritesCorrectly = testResult
    Exit Function
    
TestFail:
    Dim mockConfigErr4 As New CMockConfig
    Dim mockFileSystemErr4 As New CMockFileSystem
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogWarning_WritesToFile_Success", False
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: Initialize funciona correctamente con configuración válida y mock
Public Function Test_Initialize_WithValidConfig_Success() As CTestResult
    On Error GoTo TestFail
    
    Dim ErrorHandler As IErrorHandlerService
    Dim testResult As New CTestResult
    testResult.Initialize "Test_Initialize_WithValidConfig_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim mockConfig As New CMockConfig
    Dim mockFileSystem As New CMockFileSystem
    
    ' Arrange: Configurar el mock config
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\temp\test.log"
    
    ' Act
    errorHandlerService.Initialize mockConfig, mockFileSystem
    
    ' Assert: Si no hay error, la inicialización fue exitosa
    testResult.Pass "Initialize completado exitosamente con configuración válida y mock"
    
    Set Test_Initialize_WithValidConfig_Success = testResult
    Exit Function
    
TestFail:
    Dim mockConfigErr5 As New CMockConfig
    Dim mockFileSystemErr5 As New CMockFileSystem
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "Test_ErrorHandlerService.Test_Initialize_WithValidConfig_Success", False
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

#End If



