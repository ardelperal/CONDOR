Attribute VB_Name = "Test_LoggingService"
'******************************************************************************
' MÃ³dulo: Test_LoggingService
' PropÃ³sito: Pruebas unitarias para CLoggingService usando mocks
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' FUNCIÃ“N PRINCIPAL DE EJECUCIÃ“N
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
    
    ' Assert - Si no hay error, la inicializaciÃ³n fue exitosa
    result.Pass "El servicio se inicializÃ³ correctamente con las dependencias"
    
    GoTo Cleanup
    
ErrorHandler:
    result.Fail "Error durante la inicializaciÃ³n: " & Err.Description
    
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
    
    mockConfig.SetLogPath "C:\test\log.txt"
    service.Initialize mockConfig, mockFileSystem
    
    ' Act
    service.LogError 1001, "Error de prueba", "TestModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled Then
        If mockFileSystem.LastPath = "C:\test\log.txt" And mockFileSystem.LastMode = 8 And mockFileSystem.LastCreate = True Then
            result.Pass "LogError llamÃ³ correctamente a OpenTextFile con los parÃ¡metros esperados"
        Else
            result.Fail "LogError llamÃ³ a OpenTextFile pero con parÃ¡metros incorrectos. Path: " & mockFileSystem.LastPath & ", Mode: " & mockFileSystem.LastMode
        End If
    Else
        result.Fail "LogError no llamÃ³ a OpenTextFile"
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
    
    mockConfig.SetLogPath "C:\test\info.log"
    service.Initialize mockConfig, mockFileSystem
    
    ' Act
    service.LogInfo "Mensaje informativo", "InfoModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled Then
        If mockFileSystem.LastPath = "C:\test\info.log" And mockFileSystem.LastMode = 8 And mockFileSystem.LastCreate = True Then
            result.Pass "LogInfo llamÃ³ correctamente a OpenTextFile con los parÃ¡metros esperados"
        Else
            result.Fail "LogInfo llamÃ³ a OpenTextFile pero con parÃ¡metros incorrectos"
        End If
    Else
        result.Fail "LogInfo no llamÃ³ a OpenTextFile"
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
    
    mockConfig.SetLogPath "C:\test\warning.log"
    service.Initialize mockConfig, mockFileSystem
    
    ' Act
    service.LogWarning "Mensaje de advertencia", "WarningModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled Then
        If mockFileSystem.LastPath = "C:\test\warning.log" And mockFileSystem.LastMode = 8 And mockFileSystem.LastCreate = True Then
            result.Pass "LogWarning llamÃ³ correctamente a OpenTextFile con los parÃ¡metros esperados"
        Else
            result.Fail "LogWarning llamÃ³ a OpenTextFile pero con parÃ¡metros incorrectos"
        End If
    Else
        result.Fail "LogWarning no llamÃ³ a OpenTextFile"
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

' Prueba que LogError con todos los parÃ¡metros escribe una entrada completa
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
    
    mockConfig.SetLogPath "C:\test\complete.log"
    service.Initialize mockConfig, mockFileSystem
    
    Set mockTextFile = mockFileSystem.GetMockTextFile()
    
    ' Act
    service.LogError 500, "Error completo", "CompleteModule"
    
    ' Assert
    If mockFileSystem.WasOpenTextFileCalled And mockTextFile.WasWriteLineCalled Then
        If InStr(mockTextFile.LastWrittenLine, "ERROR") > 0 And InStr(mockTextFile.LastWrittenLine, "Error completo") > 0 Then
            result.Pass "LogError escribiÃ³ correctamente la entrada de log completa"
        Else
            result.Fail "LogError escribiÃ³ al archivo pero el contenido no es el esperado: " & mockTextFile.LastWrittenLine
        End If
    Else
        result.Fail "LogError no escribiÃ³ al archivo de log"
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
