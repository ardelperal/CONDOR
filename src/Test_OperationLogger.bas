Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_OperationLogger
' Descripción: Pruebas unitarias para el sistema de logging de operaciones.
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' ============================================================================
' PRUEBAS PARA CMockOperationLogger (Mock)
' ============================================================================

Public Function Test_MockOperationLogger_LogOperation_RecordsEntry() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger) ' Inyectar el mock
    
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Act
    logger.LogOperation "TestType", "TestEntityID", "TestDetails"
    
    ' Assert
    Call modAssert.AreEqual(1, mockLogger.LoggedOperations.Count, "Debería haber 1 operación loggeada.")
    
    Dim loggedEntry As Collection
    Set loggedEntry = mockLogger.LoggedOperations(1)
    Call modAssert.AreEqual("TestType", loggedEntry("OperationType"), "Tipo de operación incorrecto.")
    Call modAssert.AreEqual("TestEntityID", loggedEntry("EntityId"), "ID de entidad incorrecto.")
    Call modAssert.AreEqual("TestDetails", loggedEntry("Details"), "Detalles incorrectos.")
    
    Test_MockOperationLogger_LogOperation_RecordsEntry = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_MockOperationLogger_LogOperation_RecordsEntry")
    Test_MockOperationLogger_LogOperation_RecordsEntry = False
End Function

Public Function Test_MockOperationLogger_ClearLog_ResetsCount() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger)
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    logger.LogOperation "Type1", "ID1", "Details1"
    logger.LogOperation "Type2", "ID2", "Details2"
    Call modAssert.AreEqual(2, mockLogger.LoggedOperations.Count, "Debería haber 2 operaciones antes de limpiar.")
    
    ' Act
    mockLogger.ClearLog()
    
    ' Assert
    Call modAssert.AreEqual(0, mockLogger.LoggedOperations.Count, "La cuenta de operaciones debería ser 0 después de limpiar.")
    
    Test_MockOperationLogger_ClearLog_ResetsCount = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_MockOperationLogger_ClearLog_ResetsCount")
    Test_MockOperationLogger_ClearLog_ResetsCount = False
End Function

Public Function Test_OperationLoggerFactory_ReturnsMockWhenSet() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger)
    
    ' Act
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Assert
    Call modAssert.IsTrue(TypeOf logger Is CMockOperationLogger, "El factory debería devolver una instancia de CMockOperationLogger.")
    
    Test_OperationLoggerFactory_ReturnsMockWhenSet = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_OperationLoggerFactory_ReturnsMockWhenSet")
    Test_OperationLoggerFactory_ReturnsMockWhenSet = False
End Function

Public Function Test_OperationLoggerFactory_ReturnsRealWhenReset() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockLogger As New CMockOperationLogger
    Call modOperationLoggerFactory.SetMockLogger(mockLogger)
    Call modOperationLoggerFactory.ResetMockLogger()
    
    ' Act
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Assert
    Call modAssert.IsTrue(TypeOf logger Is COperationLogger, "El factory debería devolver una instancia de COperationLogger después de resetear el mock.")
    
    Test_OperationLoggerFactory_ReturnsRealWhenReset = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_OperationLoggerFactory_ReturnsRealWhenReset")
    Test_OperationLoggerFactory_ReturnsRealWhenReset = False
End Function

' ============================================================================
' PRUEBAS PARA COperationLogger (Real) - Requiere Tb_Operaciones_Log en la DB
' ============================================================================

' NOTA: Para ejecutar esta prueba, asegúrate de que la tabla 'Tb_Operaciones_Log' exista en la base de datos.
' Puedes crearla con los siguientes campos:
' ID (Autonumérico, Clave Principal)
' FechaHora (Fecha/Hora)
' Usuario (Texto corto)
' TipoOperacion (Texto corto)
' IDEntidadAfectada (Texto corto)
' Detalles (Texto largo)

Public Function Test_COperationLogger_LogOperation_WritesToDatabase() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Call modOperationLoggerFactory.ResetMockLogger() ' Asegurarse de usar el logger real
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger()
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' Limpiar logs anteriores para asegurar un conteo preciso
    db.Execute "DELETE FROM Tb_Operaciones_Log"
    
    Dim initialCount As Long
    initialCount = DCount("ID", "Tb_Operaciones_Log")
    Call modAssert.AreEqual(0, initialCount, "La tabla de logs debería estar vacía al inicio de la prueba.")
    
    ' Act
    logger.LogOperation "RealOperation", "RealEntityID", "Detalles de operación real"
    
    ' Assert
    Dim finalCount As Long
    finalCount = DCount("ID", "Tb_Operaciones_Log")
    Call modAssert.AreEqual(1, finalCount, "Debería haber 1 operación loggeada en la base de datos.")
    
    ' Opcional: Verificar el contenido del log
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM Tb_Operaciones_Log ORDER BY ID DESC", dbOpenSnapshot)
    Call modAssert.IsFalse(rs.EOF, "El recordset no debería estar vacío.")
    Call modAssert.AreEqual("RealOperation", rs!TipoOperacion, "Tipo de operación loggeado incorrecto.")
    Call modAssert.AreEqual("RealEntityID", rs!IDEntidadAfectada, "ID de entidad loggeado incorrecto.")
    Call modAssert.AreEqual("Detalles de operación real", rs!Detalles, "Detalles loggeados incorrectos.")
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Test_COperationLogger_LogOperation_WritesToDatabase = True
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_OperationLogger.Test_COperationLogger_LogOperation_WritesToDatabase")
    Test_COperationLogger_LogOperation_WritesToDatabase = False
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ============================================================================

Public Function RunOperationLoggerTests() As String
    Dim resultado As String
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    resultado = "=== PRUEBAS DE OPERATION LOGGER ===" & vbCrLf
    totalTests = 0
    passedTests = 0
    
    totalTests = totalTests + 1
    If Test_MockOperationLogger_LogOperation_RecordsEntry() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_MockOperationLogger_LogOperation_RecordsEntry" & vbCrLf
    Else
        resultado = resultado & "? Test_MockOperationLogger_LogOperation_RecordsEntry" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_MockOperationLogger_ClearLog_ResetsCount() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_MockOperationLogger_ClearLog_ResetsCount" & vbCrLf
    Else
        resultado = resultado & "? Test_MockOperationLogger_ClearLog_ResetsCount" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_OperationLoggerFactory_ReturnsMockWhenSet() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_OperationLoggerFactory_ReturnsMockWhenSet" & vbCrLf
    Else
        resultado = resultado & "? Test_OperationLoggerFactory_ReturnsMockWhenSet" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_OperationLoggerFactory_ReturnsRealWhenReset() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_OperationLoggerFactory_ReturnsRealWhenReset" & vbCrLf
    Else
        resultado = resultado & "? Test_OperationLoggerFactory_ReturnsRealWhenReset" & vbCrLf
    End If
    
    ' La siguiente prueba requiere que la tabla Tb_Operaciones_Log exista en la base de datos.
    ' Si no existe, esta prueba fallará.
    totalTests = totalTests + 1
    If Test_COperationLogger_LogOperation_WritesToDatabase() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_COperationLogger_LogOperation_WritesToDatabase" & vbCrLf
    Else
        resultado = resultado & "? Test_COperationLogger_LogOperation_WritesToDatabase" & vbCrLf
    End If
    
    resultado = resultado & vbCrLf & "Resumen: " & passedTests & "/" & totalTests & " pruebas pasadas." & vbCrLf
    
    RunOperationLoggerTests = resultado
End Function
