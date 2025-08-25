Attribute VB_Name = "Test_LoggingService"
'******************************************************************************
' Módulo: Test_LoggingService
' Propósito: Pruebas unitarias para CLoggingService usando mocks
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
'******************************************************************************

' Ejecuta todas las pruebas del LoggingService
Public Function Test_LoggingService_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "Test_LoggingService"
    
    ' Ejecutar todas las pruebas
    suite.AddTest Test_Initialize_WithValidDependencies_SetsUpCorrectly
    suite.AddTest Test_LogError_WithValidMessage_CallsFileSystemCorrectly
    suite.AddTest Test_LogInfo_WithValidMessage_CallsFileSystemCorrectly
    suite.AddTest Test_LogWarning_WithValidMessage_CallsFileSystemCorrectly
    suite.AddTest Test_LogError_WithAllParameters_WritesCompleteLogEntry
    
    Set Test_LoggingService_RunAll = suite
End Function

'******************************************************************************
' PRUEBAS UNITARIAS PRIVADAS
'******************************************************************************

' Prueba que Initialize configura correctamente las dependencias
Private Function Test_Initialize_WithValidDependencies_SetsUpCorrectly() As CTestResult
    Dim result As New CTestResult
    result.Initialize "Test_Initialize_WithValidDependencies_SetsUpCorrectly"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim service As CLoggingService
    Dim mockConfig As CMockConfig
    Dim mockFileSystem As CMockFileSystem
    
    Set service = New CLoggingService
    Set mockConfig = New CMockConfig
    Set mockFileSystem = New CMockFileSystem
    
    ' Act
    service.Initialize mockConfig, mockFileSystem
    
    ' Assert - Si no hay error, la inicialización fue exitosa
    result.Pass "El servicio se inicializó correctamente con las dependencias"
    
    GoTo Cleanup
    
ErrorHandler:
    result.Fail "Error durante la inicialización: " & Err.Description
    
Cleanup:
    Set service = Nothing
    Set mockConfig = Nothing
    Set mockFileSystem = Nothing
    Set Test_Initialize_WithValidDependencies_SetsUpCorrectly = result
End Function

' Prueba que LogError llama correctamente al sistema de ficheros
Private Function Test_LogError_WithValidMessage_CallsFileSystemCorrectly() As CTestResult
    Dim result As New CTestResult
    result.Initialize "Test_LogError_WithValidMessage_CallsFileSystemCorrectly"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim service As ILoggingService
    Dim mockConfig As CMockConfig
    Dim mockFileSystem As CMockFileSystem
    
    Set service = New CLoggingService
    Set mockConfig = New CMockConfig
    Set mockFileSystem = New CMockFileSystem
    
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\test\log.txt"
    service.Initialize mockConfig, mockFileSystem
    
    ' Act
    service.LogError 1001, "Error de prueba", "TestModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled Then
        If mockFileSystem.LastPath = "C:\test\log.txt" And mockFileSystem.LastMode = 8 And mockFileSystem.LastCreate = True Then
            result.Pass "LogError llamó correctamente a OpenTextFile con los parámetros esperados"
        Else
            result.Fail "LogError llamó a OpenTextFile pero con parámetros incorrectos. Path: " & mockFileSystem.LastPath & ", Mode: " & mockFileSystem.LastMode
        End If
    Else
        result.Fail "LogError no llamó a OpenTextFile"
    End If
    
    GoTo Cleanup
    
ErrorHandler:
    result.Fail "Error durante la prueba: " & Err.Description
    
Cleanup:
    Set service = Nothing
    Set mockConfig = Nothing
    Set mockFileSystem = Nothing
    Set Test_LogError_WithValidMessage_CallsFileSystemCorrectly = result
End Function

' Prueba que LogInfo llama correctamente al sistema de ficheros
Private Function Test_LogInfo_WithValidMessage_CallsFileSystemCorrectly() As CTestResult
    Dim result As New CTestResult
    result.Initialize "Test_LogInfo_WithValidMessage_CallsFileSystemCorrectly"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim service As ILoggingService
    Dim mockConfig As CMockConfig
    Dim mockFileSystem As CMockFileSystem
    
    Set service = New CLoggingService
    Set mockConfig = New CMockConfig
    Set mockFileSystem = New CMockFileSystem
    
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\test\info.log"
    service.Initialize mockConfig, mockFileSystem
    
    ' Act
    service.LogInfo "Mensaje informativo", "InfoModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled Then
        If mockFileSystem.LastPath = "C:\test\info.log" And mockFileSystem.LastMode = 8 And mockFileSystem.LastCreate = True Then
            result.Pass "LogInfo llamó correctamente a OpenTextFile con los parámetros esperados"
        Else
            result.Fail "LogInfo llamó a OpenTextFile pero con parámetros incorrectos"
        End If
    Else
        result.Fail "LogInfo no llamó a OpenTextFile"
    End If
    
    GoTo Cleanup
    
ErrorHandler:
    result.Fail "Error durante la prueba: " & Err.Description
    
Cleanup:
    Set service = Nothing
    Set mockConfig = Nothing
    Set mockFileSystem = Nothing
    Set Test_LogInfo_WithValidMessage_CallsFileSystemCorrectly = result
End Function

' Prueba que LogWarning llama correctamente al sistema de ficheros
Private Function Test_LogWarning_WithValidMessage_CallsFileSystemCorrectly() As CTestResult
    Dim result As New CTestResult
    result.Initialize "Test_LogWarning_WithValidMessage_CallsFileSystemCorrectly"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim service As ILoggingService
    Dim mockConfig As CMockConfig
    Dim mockFileSystem As CMockFileSystem
    
    Set service = New CLoggingService
    Set mockConfig = New CMockConfig
    Set mockFileSystem = New CMockFileSystem
    
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\test\warning.log"
    service.Initialize mockConfig, mockFileSystem
    
    ' Act
    service.LogWarning "Mensaje de advertencia", "WarningModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled Then
        If mockFileSystem.LastPath = "C:\test\warning.log" And mockFileSystem.LastMode = 8 And mockFileSystem.LastCreate = True Then
            result.Pass "LogWarning llamó correctamente a OpenTextFile con los parámetros esperados"
        Else
            result.Fail "LogWarning llamó a OpenTextFile pero con parámetros incorrectos"
        End If
    Else
        result.Fail "LogWarning no llamó a OpenTextFile"
    End If
    
    GoTo Cleanup
    
ErrorHandler:
    result.Fail "Error durante la prueba: " & Err.Description
    
Cleanup:
    Set service = Nothing
    Set mockConfig = Nothing
    Set mockFileSystem = Nothing
    Set Test_LogWarning_WithValidMessage_CallsFileSystemCorrectly = result
End Function

' Prueba que LogError con todos los parámetros escribe una entrada completa
Private Function Test_LogError_WithAllParameters_WritesCompleteLogEntry() As CTestResult
    Dim result As New CTestResult
    result.Initialize "Test_LogError_WithAllParameters_WritesCompleteLogEntry"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim service As ILoggingService
    Dim mockConfig As CMockConfig
    Dim mockFileSystem As CMockFileSystem
    Dim mockTextFile As CMockTextFile
    
    Set service = New CLoggingService
    Set mockConfig = New CMockConfig
    Set mockFileSystem = New CMockFileSystem
    
    mockConfig.AddSetting "LOG_FILE_PATH", "C:\test\complete.log"
    service.Initialize mockConfig, mockFileSystem
    
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    ' Act
    service.LogError 500, "Error completo", "CompleteModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled And mockTextFile.WasWriteLineCalled Then
        If InStr(mockTextFile.LastWrittenLine, "ERROR") > 0 And InStr(mockTextFile.LastWrittenLine, "Error completo") > 0 Then
            result.Pass "LogError escribió correctamente la entrada de log completa"
        Else
            result.Fail "LogError escribió al archivo pero el contenido no es el esperado: " & mockTextFile.LastWrittenLine
        End If
    Else
        result.Fail "LogError no escribió al archivo de log"
    End If
    
    GoTo Cleanup
    
ErrorHandler:
    result.Fail "Error durante la prueba: " & Err.Description
    
Cleanup:
    Set service = Nothing
    Set mockConfig = Nothing
    Set mockFileSystem = Nothing
    Set mockTextFile = Nothing
    Set Test_LogError_WithAllParameters_WritesCompleteLogEntry = result
End Function
