Attribute VB_Name = "Test_ErrorHandlerService"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' =====================================================
' Test_ErrorHandlerService.bas
' MÃ³dulo de pruebas para CErrorHandlerService
' Implementa pruebas unitarias aisladas usando mocks
' =====================================================

' FunciÃ³n principal para ejecutar todas las pruebas del ErrorHandlerService
Public Function Test_ErrorHandlerService_RunAll() As CTestSuiteResult
    On Error GoTo ErrorHandler
    
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_ErrorHandlerService"
    
    ' Ejecutar todas las pruebas
    Call suiteResult.AddTestResult(Test_LogError_WritesToFile_Success())
    Call suiteResult.AddTestResult(Test_LogInfo_WritesToFile_Success())
    Call suiteResult.AddTestResult(Test_LogWarning_WritesToFile_Success())
    Call suiteResult.AddTestResult(Test_Initialize_WithValidConfig_Success())
    
    Set Test_ErrorHandlerService_RunAll = suiteResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Call errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_ErrorHandlerService_RunAll")
    Set Test_ErrorHandlerService_RunAll = Nothing
End Function

' Prueba principal: LogError escribe correctamente usando mock
Public Function Test_LogError_WritesToFile_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogError_WritesToFile_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockFileSystem As New CMockFileSystem
    Dim mockTextFile As CMockTextFile
    Dim writtenContent As String
    
    ' Arrange: Configurar el mock y el servicio
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    ' Inicializar el servicio con el mock FileSystem
    errorHandlerService.Initialize configService, mockFileSystem
    
    ' Act: Llamar al mÃ©todo LogError
    errorHandlerService.LogError 1001, "Error de prueba", "Test_Module.Test_Function"
    
    ' Assert: Verificar que el mock capturÃ³ el contenido esperado
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "1001") > 0, "El nÃºmero de error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "Error de prueba") > 0, "La descripciÃ³n del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "Test_Module.Test_Function") > 0, "La fuente del error debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "ERROR") > 0, "El nivel ERROR debe estar en el log")
    
    testResult.SetResult True, "LogError escribiÃ³ correctamente usando mock"
    
    Set Test_LogError_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Call errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogError_WritesToFile_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: LogInfo escribe correctamente usando mock
Public Function Test_LogInfo_WritesToFile_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogInfo_WritesToFile_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockFileSystem As New CMockFileSystem
    Dim mockTextFile As CMockTextFile
    Dim writtenContent As String
    
    ' Arrange
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    errorHandlerService.Initialize configService, mockFileSystem
    
    ' Act
    errorHandlerService.LogInfo "InformaciÃ³n de prueba", "Test_Module.Test_Function"
    
    ' Assert
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "InformaciÃ³n de prueba") > 0, "El mensaje de info debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "INFO") > 0, "El nivel INFO debe estar en el log")
    
    testResult.SetResult True, "LogInfo escribiÃ³ correctamente usando mock"
    
    Set Test_LogInfo_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Call errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogInfo_WritesToFile_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: LogWarning escribe correctamente usando mock
Public Function Test_LogWarning_WritesToFile_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_LogWarning_WritesToFile_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockFileSystem As New CMockFileSystem
    Dim mockTextFile As CMockTextFile
    Dim writtenContent As String
    
    ' Arrange
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    errorHandlerService.Initialize configService, mockFileSystem
    
    ' Act
    errorHandlerService.LogWarning "Advertencia de prueba", "Test_Module.Test_Function"
    
    ' Assert
    writtenContent = mockTextFile.LastWrittenLine
    
    Call modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    Call modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    Call modAssert.IsTrue(InStr(writtenContent, "Advertencia de prueba") > 0, "El mensaje de warning debe estar en el log")
    Call modAssert.IsTrue(InStr(writtenContent, "WARNING") > 0, "El nivel WARNING debe estar en el log")
    
    testResult.SetResult True, "LogWarning escribiÃ³ correctamente usando mock"
    
    Set Test_LogWarning_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Call errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogWarning_WritesToFile_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: Initialize funciona correctamente con configuraciÃ³n vÃ¡lida y mock
Public Function Test_Initialize_WithValidConfig_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.Initialize "Test_Initialize_WithValidConfig_Success"
    
    ' Variables para la prueba
    Dim errorHandlerService As New CErrorHandlerService
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockFileSystem As New CMockFileSystem
    
    ' Arrange
    ' Las pruebas usan el mock del sistema de ficheros para aislamiento completo
    
    ' Act
    errorHandlerService.Initialize configService, mockFileSystem
    
    ' Assert: Si no hay error, la inicializaciÃ³n fue exitosa
    testResult.SetResult True, "Initialize completado exitosamente con configuraciÃ³n vÃ¡lida y mock"
    
    Set Test_Initialize_WithValidConfig_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Call errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_Initialize_WithValidConfig_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

#End If
