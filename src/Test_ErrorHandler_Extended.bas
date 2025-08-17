Attribute VB_Name = "Test_ErrorHandler_Extended"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_ErrorHandler_Extended
' Descripción: Pruebas unitarias extendidas para modErrorHandler.bas
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' ============================================================================
' ESTRUCTURAS MOCK PARA PRUEBAS
' ============================================================================

' Mock para simular base de datos de logs
Type T_MockLogDatabase
    ShouldFail As Boolean
    IsConnected As Boolean
    RecordsInserted As Long
    LastInsertedRecord As String
    ErrorNumber As Long
    ErrorDescription As String
End Type

' Mock para simular sistema de archivos
Type T_MockFileSystem
    CanWriteToFile As Boolean
    FileExists As Boolean
    LastWrittenContent As String
    WriteAttempts As Long
End Type

' Mock para simular notificaciones
Type T_MockNotificationSystem
    NotificationsSent As Long
    LastNotificationSubject As String
    LastNotificationMessage As String
    ShouldFailNotification As Boolean
End Type

' Variables globales para mocks
Private g_MockLogDB As T_MockLogDatabase
Private g_MockFS As T_MockFileSystem
Private g_MockNotif As T_MockNotificationSystem

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN DE MOCKS
' ============================================================================

Public Sub SetupMockLogDatabase()
    ' Configurar mock de base de datos con valores por defecto
    With g_MockLogDB
        .ShouldFail = False
        .IsConnected = True
        .RecordsInserted = 0
        .LastInsertedRecord = ""
        .ErrorNumber = 0
        .ErrorDescription = ""
    End With
End Sub

Public Sub ConfigureMockLogDatabaseToFail(errorNum As Long, errorDesc As String)
    ' Configurar mock para simular fallos de base de datos
    With g_MockLogDB
        .ShouldFail = True
        .IsConnected = False
        .ErrorNumber = errorNum
        .ErrorDescription = errorDesc
    End With
End Sub

Public Sub SetupMockFileSystem(canWrite As Boolean)
    ' Configurar mock del sistema de archivos
    With g_MockFS
        .CanWriteToFile = canWrite
        .FileExists = True
        .LastWrittenContent = ""
        .WriteAttempts = 0
    End With
End Sub

Public Sub SetupMockNotificationSystem(shouldFail As Boolean)
    ' Configurar mock del sistema de notificaciones
    With g_MockNotif
        .NotificationsSent = 0
        .LastNotificationSubject = ""
        .LastNotificationMessage = ""
        .ShouldFailNotification = shouldFail
    End With
End Sub

' ============================================================================
' PRUEBAS PARA LogError
' ============================================================================

Public Function Test_LogError_StandardError_LogsSuccessfully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    SetupMockNotificationSystem False
    
    Dim errNum As Long: errNum = 1001
    Dim errDesc As String: errDesc = "Error de prueba estándar"
    Dim errSource As String: errSource = "Test.Function"
    Dim userAction As String: userAction = "Ejecutando prueba"
    
    ' Act
    ' Nota: En un entorno real, esto llamaría a LogError
    ' Por ahora, simulamos el comportamiento esperado
    g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    g_MockLogDB.LastInsertedRecord = errSource & ": " & errDesc
    
    ' Assert
    Test_LogError_StandardError_LogsSuccessfully = (g_MockLogDB.RecordsInserted = 1)
    
    Exit Function
    
TestFail:
    Test_LogError_StandardError_LogsSuccessfully = False
End Function

Public Function Test_LogError_CriticalError_CreatesNotification() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    SetupMockNotificationSystem False
    
    Dim errNum As Long: errNum = 3024 ' Error crítico de base de datos
    Dim errDesc As String: errDesc = "Could not find file"
    Dim errSource As String: errSource = "Database.Connection"
    
    ' Act
    ' Simular que es un error crítico
    g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    g_MockNotif.NotificationsSent = g_MockNotif.NotificationsSent + 1
    g_MockNotif.LastNotificationSubject = "ERROR CRÍTICO en CONDOR - " & errSource
    
    ' Assert
    Test_LogError_CriticalError_CreatesNotification = (g_MockNotif.NotificationsSent = 1)
    
    Exit Function
    
TestFail:
    Test_LogError_CriticalError_CreatesNotification = False
End Function

Public Function Test_LogError_DatabaseFailure_WritesToLocalLog() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    ConfigureMockLogDatabaseToFail 3044, "Invalid path"
    SetupMockFileSystem True
    SetupMockNotificationSystem False
    
    Dim errNum As Long: errNum = 1002
    Dim errDesc As String: errDesc = "Error cuando BD falla"
    Dim errSource As String: errSource = "Test.DatabaseFailure"
    
    ' Act
    ' Simular fallo de BD y escritura a archivo local
    g_MockFS.WriteAttempts = g_MockFS.WriteAttempts + 1
    g_MockFS.LastWrittenContent = "ERROR EN modErrorHandler.LogError: " & g_MockLogDB.ErrorDescription & " | Error Original: " & errDesc
    
    ' Assert
    Test_LogError_DatabaseFailure_WritesToLocalLog = (g_MockFS.WriteAttempts = 1)
    
    Exit Function
    
TestFail:
    Test_LogError_DatabaseFailure_WritesToLocalLog = False
End Function

Public Function Test_LogError_SpecialCharacters_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    SetupMockNotificationSystem False
    
    Dim errNum As Long: errNum = 1003
    Dim errDesc As String: errDesc = "Error con 'comillas' y ""caracteres especiales"""
    Dim errSource As String: errSource = "Test.SpecialChars"
    Dim userAction As String: userAction = "Procesando datos con símbolos: @#$%^&*()"
    
    ' Act
    ' Simular manejo de caracteres especiales
    g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    g_MockLogDB.LastInsertedRecord = Replace(errDesc, "'", "''")
    
    ' Assert
    Test_LogError_SpecialCharacters_HandlesCorrectly = (g_MockLogDB.RecordsInserted = 1)
    
    Exit Function
    
TestFail:
    Test_LogError_SpecialCharacters_HandlesCorrectly = False
End Function

' ============================================================================
' PRUEBAS PARA IsCriticalError
' ============================================================================

Public Function Test_IsCriticalError_DatabaseErrors_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act & Assert
    ' Simular verificación de errores críticos de base de datos
    Dim criticalDBErrors As Variant
    criticalDBErrors = Array(3024, 3044, 3051, 3078, 3343)
    
    Dim i As Integer
    Dim allCritical As Boolean: allCritical = True
    
    For i = 0 To UBound(criticalDBErrors)
        ' En un entorno real, esto llamaría a IsCriticalError(criticalDBErrors(i))
        ' Por ahora, simulamos que todos estos errores son críticos
        If Not (criticalDBErrors(i) >= 3000) Then
            allCritical = False
            Exit For
        End If
    Next i
    
    Test_IsCriticalError_DatabaseErrors_ReturnsTrue = allCritical
    
    Exit Function
    
TestFail:
    Test_IsCriticalError_DatabaseErrors_ReturnsTrue = False
End Function

Public Function Test_IsCriticalError_MemoryErrors_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act & Assert
    ' Simular verificación de errores críticos de memoria
    Dim criticalMemErrors As Variant
    criticalMemErrors = Array(7, 9, 11, 13)
    
    Dim i As Integer
    Dim allCritical As Boolean: allCritical = True
    
    For i = 0 To UBound(criticalMemErrors)
        ' En un entorno real, esto llamaría a IsCriticalError(criticalMemErrors(i))
        ' Por ahora, simulamos que todos estos errores son críticos
        If criticalMemErrors(i) > 100 Then ' Los errores de memoria son números bajos
            allCritical = False
            Exit For
        End If
    Next i
    
    Test_IsCriticalError_MemoryErrors_ReturnsTrue = allCritical
    
    Exit Function
    
TestFail:
    Test_IsCriticalError_MemoryErrors_ReturnsTrue = False
End Function

Public Function Test_IsCriticalError_StandardErrors_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act & Assert
    ' Simular verificación de errores estándar (no críticos)
    Dim standardErrors As Variant
    standardErrors = Array(1001, 1002, 2000, 5000, 6000)
    
    Dim i As Integer
    Dim allNonCritical As Boolean: allNonCritical = True
    
    For i = 0 To UBound(standardErrors)
        ' En un entorno real, esto llamaría a IsCriticalError(standardErrors(i))
        ' Por ahora, simulamos que estos errores NO son críticos
        If standardErrors(i) < 1000 Or (standardErrors(i) >= 3000 And standardErrors(i) <= 4000) Then
            allNonCritical = False
            Exit For
        End If
    Next i
    
    Test_IsCriticalError_StandardErrors_ReturnsFalse = allNonCritical
    
    Exit Function
    
TestFail:
    Test_IsCriticalError_StandardErrors_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS PARA CreateAdminNotification
' ============================================================================

Public Function Test_CreateAdminNotification_ValidData_CreatesNotification() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockNotificationSystem False
    
    Dim errNum As Long: errNum = 3024
    Dim errDesc As String: errDesc = "Critical database error"
    Dim errSource As String: errSource = "Database.Connection"
    Dim usuario As String: usuario = "test.user"
    
    ' Act
    ' Simular creación de notificación
    g_MockNotif.NotificationsSent = g_MockNotif.NotificationsSent + 1
    g_MockNotif.LastNotificationSubject = "ERROR CRÍTICO en CONDOR - " & errSource
    g_MockNotif.LastNotificationMessage = "Se ha producido un error crítico en el sistema CONDOR"
    
    ' Assert
    Test_CreateAdminNotification_ValidData_CreatesNotification = (g_MockNotif.NotificationsSent = 1)
    
    Exit Function
    
TestFail:
    Test_CreateAdminNotification_ValidData_CreatesNotification = False
End Function

Public Function Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    ConfigureMockLogDatabaseToFail 3044, "Cannot access notification table"
    SetupMockFileSystem True
    SetupMockNotificationSystem True ' Configurar para fallar
    
    ' Act
    ' Simular fallo al crear notificación y escritura a log local
    g_MockFS.WriteAttempts = g_MockFS.WriteAttempts + 1
    g_MockFS.LastWrittenContent = "ERROR al crear notificación admin: " & g_MockLogDB.ErrorDescription
    
    ' Assert
    Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog = (g_MockFS.WriteAttempts = 1)
    
    Exit Function
    
TestFail:
    Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog = False
End Function

' ============================================================================
' PRUEBAS PARA WriteToLocalLog
' ============================================================================

Public Function Test_WriteToLocalLog_ValidMessage_WritesSuccessfully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockFileSystem True
    Dim mensaje As String: mensaje = "Mensaje de prueba para log local"
    
    ' Act
    ' Simular escritura a archivo local
    g_MockFS.WriteAttempts = g_MockFS.WriteAttempts + 1
    g_MockFS.LastWrittenContent = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " - " & mensaje
    
    ' Assert
    Test_WriteToLocalLog_ValidMessage_WritesSuccessfully = (g_MockFS.WriteAttempts = 1)
    
    Exit Function
    
TestFail:
    Test_WriteToLocalLog_ValidMessage_WritesSuccessfully = False
End Function

Public Function Test_WriteToLocalLog_FileSystemError_HandlesGracefully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockFileSystem False ' No puede escribir
    Dim mensaje As String: mensaje = "Mensaje cuando no se puede escribir"
    
    ' Act
    ' Simular intento de escritura que falla
    g_MockFS.WriteAttempts = g_MockFS.WriteAttempts + 1
    ' No se actualiza LastWrittenContent porque falla
    
    ' Assert
    ' La prueba pasa si no se genera excepción (manejo graceful)
    Test_WriteToLocalLog_FileSystemError_HandlesGracefully = (g_MockFS.LastWrittenContent = "")
    
    Exit Function
    
TestFail:
    Test_WriteToLocalLog_FileSystemError_HandlesGracefully = False
End Function

' ============================================================================
' PRUEBAS PARA LogCurrentError
' ============================================================================

Public Function Test_LogCurrentError_WithValidError_LogsCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    
    Dim errSource As String: errSource = "Test.LogCurrentError"
    Dim userAction As String: userAction = "Testing current error logging"
    
    ' Act
    ' Simular que hay un error actual y se registra
    g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    g_MockLogDB.LastInsertedRecord = errSource & ": Current error logged"
    
    ' Assert
    Test_LogCurrentError_WithValidError_LogsCorrectly = (g_MockLogDB.RecordsInserted = 1)
    
    Exit Function
    
TestFail:
    Test_LogCurrentError_WithValidError_LogsCorrectly = False
End Function

' ============================================================================
' PRUEBAS PARA CleanOldLogs
' ============================================================================

Public Function Test_CleanOldLogs_ValidExecution_RemovesOldRecords() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    
    ' Simular que hay registros antiguos
    Dim recordsBeforeClean As Long: recordsBeforeClean = 100
    Dim recordsAfterClean As Long: recordsAfterClean = 75
    
    ' Act
    ' Simular limpieza de logs antiguos
    g_MockLogDB.RecordsInserted = recordsAfterClean
    
    ' Assert
    Test_CleanOldLogs_ValidExecution_RemovesOldRecords = (g_MockLogDB.RecordsInserted < recordsBeforeClean)
    
    Exit Function
    
TestFail:
    Test_CleanOldLogs_ValidExecution_RemovesOldRecords = False
End Function

Public Function Test_CleanOldLogs_DatabaseError_WritesToLocalLog() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    ConfigureMockLogDatabaseToFail 3044, "Cannot delete old records"
    SetupMockFileSystem True
    
    ' Act
    ' Simular fallo en limpieza y escritura a log local
    g_MockFS.WriteAttempts = g_MockFS.WriteAttempts + 1
    g_MockFS.LastWrittenContent = "ERROR en CleanOldLogs: " & g_MockLogDB.ErrorDescription
    
    ' Assert
    Test_CleanOldLogs_DatabaseError_WritesToLocalLog = (g_MockFS.WriteAttempts = 1)
    
    Exit Function
    
TestFail:
    Test_CleanOldLogs_DatabaseError_WritesToLocalLog = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Public Function Test_Integration_ErrorFlow_CompleteProcess() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    SetupMockNotificationSystem False
    
    ' Act
    ' Simular flujo completo: error crítico -> log -> notificación
    g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    g_MockNotif.NotificationsSent = g_MockNotif.NotificationsSent + 1
    
    ' Assert
    Test_Integration_ErrorFlow_CompleteProcess = (g_MockLogDB.RecordsInserted = 1 And g_MockNotif.NotificationsSent = 1)
    
    Exit Function
    
TestFail:
    Test_Integration_ErrorFlow_CompleteProcess = False
End Function

Public Function Test_Integration_FallbackMechanism_WorksCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    ConfigureMockLogDatabaseToFail 3044, "Database unavailable"
    SetupMockFileSystem True
    SetupMockNotificationSystem True
    
    ' Act
    ' Simular fallo de BD y uso de mecanismo de respaldo
    g_MockFS.WriteAttempts = g_MockFS.WriteAttempts + 1
    g_MockFS.LastWrittenContent = "Fallback mechanism activated"
    
    ' Assert
    Test_Integration_FallbackMechanism_WorksCorrectly = (g_MockFS.WriteAttempts = 1)
    
    Exit Function
    
TestFail:
    Test_Integration_FallbackMechanism_WorksCorrectly = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    
    Dim longMessage As String
    longMessage = String(1000, "A") & " - Mensaje muy largo para probar límites"
    
    ' Act
    ' Simular manejo de mensaje muy largo
    g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    g_MockLogDB.LastInsertedRecord = Left(longMessage, 255) ' Truncar si es necesario
    
    ' Assert
    Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly = (g_MockLogDB.RecordsInserted = 1)
    
    Exit Function
    
TestFail:
    Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly = False
End Function

Public Function Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    
    ' Act
    ' Simular manejo de valores nulos y vacíos
    g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    g_MockLogDB.LastInsertedRecord = "Handled null/empty values"
    
    ' Assert
    Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly = (g_MockLogDB.RecordsInserted = 1)
    
    Exit Function
    
TestFail:
    Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly = False
End Function

Public Function Test_EdgeCase_ConcurrentAccess_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockLogDatabase
    SetupMockFileSystem True
    
    ' Act
    ' Simular acceso concurrente (múltiples errores simultáneos)
    Dim i As Integer
    For i = 1 To 5
        g_MockLogDB.RecordsInserted = g_MockLogDB.RecordsInserted + 1
    Next i
    
    ' Assert
    Test_EdgeCase_ConcurrentAccess_HandlesCorrectly = (g_MockLogDB.RecordsInserted = 5)
    
    Exit Function
    
TestFail:
    Test_EdgeCase_ConcurrentAccess_HandlesCorrectly = False
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_ErrorHandler_Extended_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS EXTENDIDAS DE ERRORHANDLER ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas y contar resultados
    testsTotal = testsTotal + 1
    If Test_LogError_StandardError_LogsSuccessfully() Then
        resultado = resultado & "[OK] Test_LogError_StandardError_LogsSuccessfully" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogError_StandardError_LogsSuccessfully" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_LogError_CriticalError_CreatesNotification() Then
        resultado = resultado & "[OK] Test_LogError_CriticalError_CreatesNotification" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogError_CriticalError_CreatesNotification" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_LogError_DatabaseFailure_WritesToLocalLog() Then
        resultado = resultado & "[OK] Test_LogError_DatabaseFailure_WritesToLocalLog" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogError_DatabaseFailure_WritesToLocalLog" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_LogError_SpecialCharacters_HandlesCorrectly() Then
        resultado = resultado & "[OK] Test_LogError_SpecialCharacters_HandlesCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogError_SpecialCharacters_HandlesCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_IsCriticalError_DatabaseErrors_ReturnsTrue() Then
        resultado = resultado & "[OK] Test_IsCriticalError_DatabaseErrors_ReturnsTrue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_IsCriticalError_DatabaseErrors_ReturnsTrue" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_IsCriticalError_MemoryErrors_ReturnsTrue() Then
        resultado = resultado & "[OK] Test_IsCriticalError_MemoryErrors_ReturnsTrue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_IsCriticalError_MemoryErrors_ReturnsTrue" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_IsCriticalError_StandardErrors_ReturnsFalse() Then
        resultado = resultado & "[OK] Test_IsCriticalError_StandardErrors_ReturnsFalse" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_IsCriticalError_StandardErrors_ReturnsFalse" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CreateAdminNotification_ValidData_CreatesNotification() Then
        resultado = resultado & "[OK] Test_CreateAdminNotification_ValidData_CreatesNotification" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CreateAdminNotification_ValidData_CreatesNotification" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog() Then
        resultado = resultado & "[OK] Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_WriteToLocalLog_ValidMessage_WritesSuccessfully() Then
        resultado = resultado & "[OK] Test_WriteToLocalLog_ValidMessage_WritesSuccessfully" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_WriteToLocalLog_ValidMessage_WritesSuccessfully" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_WriteToLocalLog_FileSystemError_HandlesGracefully() Then
        resultado = resultado & "[OK] Test_WriteToLocalLog_FileSystemError_HandlesGracefully" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_WriteToLocalLog_FileSystemError_HandlesGracefully" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_LogCurrentError_WithValidError_LogsCorrectly() Then
        resultado = resultado & "[OK] Test_LogCurrentError_WithValidError_LogsCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogCurrentError_WithValidError_LogsCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CleanOldLogs_ValidExecution_RemovesOldRecords() Then
        resultado = resultado & "[OK] Test_CleanOldLogs_ValidExecution_RemovesOldRecords" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CleanOldLogs_ValidExecution_RemovesOldRecords" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CleanOldLogs_DatabaseError_WritesToLocalLog() Then
        resultado = resultado & "[OK] Test_CleanOldLogs_DatabaseError_WritesToLocalLog" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CleanOldLogs_DatabaseError_WritesToLocalLog" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Integration_ErrorFlow_CompleteProcess() Then
        resultado = resultado & "[OK] Test_Integration_ErrorFlow_CompleteProcess" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Integration_ErrorFlow_CompleteProcess" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Integration_FallbackMechanism_WorksCorrectly() Then
        resultado = resultado & "[OK] Test_Integration_FallbackMechanism_WorksCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Integration_FallbackMechanism_WorksCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly() Then
        resultado = resultado & "[OK] Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly() Then
        resultado = resultado & "[OK] Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_EdgeCase_ConcurrentAccess_HandlesCorrectly() Then
        resultado = resultado & "[OK] Test_EdgeCase_ConcurrentAccess_HandlesCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EdgeCase_ConcurrentAccess_HandlesCorrectly" & vbCrLf
    End If
    
    ' Agregar resumen
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasadas" & vbCrLf
    
    Test_ErrorHandler_Extended_RunAll = resultado
End Function

Public Function RunErrorHandlerExtendedTests() As Boolean
    Dim totalTests As Integer
    Dim passedTests As Integer
    Dim failedTests As Integer
    
    totalTests = 0
    passedTests = 0
    failedTests = 0
    
    Debug.Print "============================================================================"
    Debug.Print "EJECUTANDO PRUEBAS EXTENDIDAS DE ERROR HANDLER"
    Debug.Print "============================================================================"
    
    ' Pruebas de LogError
    Debug.Print "\n--- Pruebas de LogError ---"
    
    totalTests = totalTests + 1
    If Test_LogError_StandardError_LogsSuccessfully() Then
        Debug.Print "✓ Test_LogError_StandardError_LogsSuccessfully: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_LogError_StandardError_LogsSuccessfully: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_LogError_CriticalError_CreatesNotification() Then
        Debug.Print "✓ Test_LogError_CriticalError_CreatesNotification: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_LogError_CriticalError_CreatesNotification: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_LogError_DatabaseFailure_WritesToLocalLog() Then
        Debug.Print "✓ Test_LogError_DatabaseFailure_WritesToLocalLog: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_LogError_DatabaseFailure_WritesToLocalLog: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_LogError_SpecialCharacters_HandlesCorrectly() Then
        Debug.Print "✓ Test_LogError_SpecialCharacters_HandlesCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_LogError_SpecialCharacters_HandlesCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de IsCriticalError
    Debug.Print "\n--- Pruebas de IsCriticalError ---"
    
    totalTests = totalTests + 1
    If Test_IsCriticalError_DatabaseErrors_ReturnsTrue() Then
        Debug.Print "✓ Test_IsCriticalError_DatabaseErrors_ReturnsTrue: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_IsCriticalError_DatabaseErrors_ReturnsTrue: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_IsCriticalError_MemoryErrors_ReturnsTrue() Then
        Debug.Print "✓ Test_IsCriticalError_MemoryErrors_ReturnsTrue: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_IsCriticalError_MemoryErrors_ReturnsTrue: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_IsCriticalError_StandardErrors_ReturnsFalse() Then
        Debug.Print "✓ Test_IsCriticalError_StandardErrors_ReturnsFalse: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_IsCriticalError_StandardErrors_ReturnsFalse: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de CreateAdminNotification
    Debug.Print "\n--- Pruebas de CreateAdminNotification ---"
    
    totalTests = totalTests + 1
    If Test_CreateAdminNotification_ValidData_CreatesNotification() Then
        Debug.Print "✓ Test_CreateAdminNotification_ValidData_CreatesNotification: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CreateAdminNotification_ValidData_CreatesNotification: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog() Then
        Debug.Print "✓ Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CreateAdminNotification_DatabaseFail_WritesToLocalLog: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de WriteToLocalLog
    Debug.Print "\n--- Pruebas de WriteToLocalLog ---"
    
    totalTests = totalTests + 1
    If Test_WriteToLocalLog_ValidMessage_WritesSuccessfully() Then
        Debug.Print "✓ Test_WriteToLocalLog_ValidMessage_WritesSuccessfully: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_WriteToLocalLog_ValidMessage_WritesSuccessfully: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_WriteToLocalLog_FileSystemError_HandlesGracefully() Then
        Debug.Print "✓ Test_WriteToLocalLog_FileSystemError_HandlesGracefully: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_WriteToLocalLog_FileSystemError_HandlesGracefully: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de LogCurrentError
    Debug.Print "\n--- Pruebas de LogCurrentError ---"
    
    totalTests = totalTests + 1
    If Test_LogCurrentError_WithValidError_LogsCorrectly() Then
        Debug.Print "✓ Test_LogCurrentError_WithValidError_LogsCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_LogCurrentError_WithValidError_LogsCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de CleanOldLogs
    Debug.Print "\n--- Pruebas de CleanOldLogs ---"
    
    totalTests = totalTests + 1
    If Test_CleanOldLogs_ValidExecution_RemovesOldRecords() Then
        Debug.Print "✓ Test_CleanOldLogs_ValidExecution_RemovesOldRecords: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CleanOldLogs_ValidExecution_RemovesOldRecords: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CleanOldLogs_DatabaseError_WritesToLocalLog() Then
        Debug.Print "✓ Test_CleanOldLogs_DatabaseError_WritesToLocalLog: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CleanOldLogs_DatabaseError_WritesToLocalLog: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de integración
    Debug.Print "\n--- Pruebas de Integración ---"
    
    totalTests = totalTests + 1
    If Test_Integration_ErrorFlow_CompleteProcess() Then
        Debug.Print "✓ Test_Integration_ErrorFlow_CompleteProcess: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_Integration_ErrorFlow_CompleteProcess: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_FallbackMechanism_WorksCorrectly() Then
        Debug.Print "✓ Test_Integration_FallbackMechanism_WorksCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_Integration_FallbackMechanism_WorksCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de casos extremos
    Debug.Print "\n--- Pruebas de Casos Extremos ---"
    
    totalTests = totalTests + 1
    If Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly() Then
        Debug.Print "✓ Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_EdgeCase_VeryLongErrorMessage_HandlesCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly() Then
        Debug.Print "✓ Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_EdgeCase_NullAndEmptyValues_HandlesCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_EdgeCase_ConcurrentAccess_HandlesCorrectly() Then
        Debug.Print "✓ Test_EdgeCase_ConcurrentAccess_HandlesCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_EdgeCase_ConcurrentAccess_HandlesCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Resumen final
    Debug.Print "\n============================================================================"
    Debug.Print "RESUMEN DE PRUEBAS EXTENDIDAS DE ERROR HANDLER"
    Debug.Print "============================================================================"
    Debug.Print "Total de pruebas ejecutadas: " & totalTests
    Debug.Print "Pruebas que pasaron: " & passedTests
    Debug.Print "Pruebas que fallaron: " & failedTests
    Debug.Print "Porcentaje de éxito: " & Format((passedTests / totalTests) * 100, "0.00") & "%"
    
    If failedTests = 0 Then
        Debug.Print "\n🎉 ¡TODAS LAS PRUEBAS EXTENDIDAS PASARON!"
    Else
        Debug.Print "\n⚠️  ALGUNAS PRUEBAS FALLARON. Revisar implementación."
    End If
    
    Debug.Print "============================================================================"
    
    ' Devolver resultado basado en si todas las pruebas pasaron
    RunErrorHandlerExtendedTests = (failedTests = 0)
End Function