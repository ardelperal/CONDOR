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
Public Function Test_ErrorHandlerService_RunAll() As CTestSuiteResult
    On Error GoTo ErrorHandler
    
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_ErrorHandlerService"
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult(Test_LogError_WritesToFile_Success())
    suiteResult.AddTestResult(Test_LogInfo_WritesToFile_Success())
    suiteResult.AddTestResult(Test_LogWarning_WritesToFile_Success())
    suiteResult.AddTestResult(Test_Initialize_WithValidConfig_Success())
    
    Set Test_ErrorHandlerService_RunAll = suiteResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_ErrorHandlerService_RunAll")
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
    
    ' Act: Llamar al método LogError
    errorHandlerService.LogError 1001, "Error de prueba", "Test_Module.Test_Function"
    
    ' Assert: Verificar que el mock capturó el contenido esperado
    writtenContent = mockTextFile.LastWrittenLine
    
    modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    modAssert.IsTrue(InStr(writtenContent, "1001") > 0, "El número de error debe estar en el log")
    modAssert.IsTrue(InStr(writtenContent, "Error de prueba") > 0, "La descripción del error debe estar en el log")
    modAssert.IsTrue(InStr(writtenContent, "Test_Module.Test_Function") > 0, "La fuente del error debe estar en el log")
    modAssert.IsTrue(InStr(writtenContent, "ERROR") > 0, "El nivel ERROR debe estar en el log")
    
    testResult.Pass "LogError escribió correctamente usando mock"
    
    Set Test_LogError_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogError_WritesToFile_Success")
    testResult.Fail "Error durante la prueba: " & Err.Description
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
    errorHandlerService.LogInfo "Información de prueba", "Test_Module.Test_Function"
    
    ' Assert
    writtenContent = mockTextFile.LastWrittenLine
    
    modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    modAssert.IsTrue(InStr(writtenContent, "Información de prueba") > 0, "El mensaje de info debe estar en el log")
    modAssert.IsTrue(InStr(writtenContent, "INFO") > 0, "El nivel INFO debe estar en el log")
    
    testResult.Pass "LogInfo escribió correctamente usando mock"
    
    Set Test_LogInfo_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogInfo_WritesToFile_Success")
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
    
    modAssert.IsTrue(mockFileSystem.WasOpenTextFileCalled, "OpenTextFile debe haber sido llamado")
    modAssert.IsTrue(mockTextFile.WasWriteLineCalled, "WriteLine debe haber sido llamado")
    modAssert.IsTrue(mockTextFile.WasCloseCalled, "Close debe haber sido llamado")
    modAssert.IsTrue(InStr(writtenContent, "Advertencia de prueba") > 0, "El mensaje de warning debe estar en el log")
    modAssert.IsTrue(InStr(writtenContent, "WARNING") > 0, "El nivel WARNING debe estar en el log")
    
    testResult.Pass "LogWarning escribió correctamente usando mock"
    
    Set Test_LogWarning_WritesToFile_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_LogWarning_WritesToFile_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

' Prueba: Initialize funciona correctamente con configuración válida y mock
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
    
    ' Assert: Si no hay error, la inicialización fue exitosa
    testResult.Pass "Initialize completado exitosamente con configuración válida y mock"
    
    Set Test_Initialize_WithValidConfig_Success = testResult
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandlerService.Test_Initialize_WithValidConfig_Success")
    testResult.SetResult False, "Error durante la prueba: " & Err.Description
End Function

#End If