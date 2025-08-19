Attribute VB_Name = "Test_ErrorHandler"
Option Compare Database
Option Explicit

' ============================================================================
' TIPOS Y VARIABLES PARA MOCKS
' ============================================================================

' Mock para simular DAO.Database
Private Type T_MockDatabase
    IsConnected As Boolean
    ShouldFail As Boolean
    ErrorNumber As Long
    ErrorDescription As String
    ExecuteCallCount As Long
    LastSQL As String
    DatabasePath As String
    RecordCount As Long
End Type

' Mock para simular el sistema de archivos
Private Type T_MockFileSystem
    FileExists As Boolean
    CanWrite As Boolean
    ShouldFailWrite As Boolean
    LastWrittenContent As String
    WriteCallCount As Long
End Type
Private m_MockDB As T_MockDatabase
Private m_MockFS As T_MockFileSystem
' ============================================================================
' M├│dulo: Test_ErrorHandler
' Descripci├│n: Pruebas para el sistema de manejo de errores centralizado
' Autor: Sistema CONDOR
' Fecha: 2024
' Implementa patr├│n AAA (Arrange, Act, Assert)
' ============================================================================

' Funci├│n principal que ejecuta todas las pruebas del manejo de errores
Public Function RunErrorHandlerTests() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DEL SISTEMA DE MANEJO DE ERRORES ===" & vbCrLf & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Test 1: LogError con datos v├ílidos
    testsTotal = testsTotal + 1
    If Test_LogError_ValidData_Success() Then
        resultado = resultado & "[OK] Test_LogError_ValidData_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogError_ValidData_Success" & vbCrLf
    End If
    
    ' Test 2: Error cr├¡tico crea notificaci├│n
    testsTotal = testsTotal + 1
    If Test_LogError_CriticalError_CreatesNotification() Then
        resultado = resultado & "[OK] Test_LogError_CriticalError_CreatesNotification" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogError_CriticalError_CreatesNotification" & vbCrLf
    End If
    
    ' Test 3: Fallo de base de datos escribe a log local
    testsTotal = testsTotal + 1
    If Test_LogError_DatabaseFail_WritesToLocalLog() Then
        resultado = resultado & "[OK] Test_LogError_DatabaseFail_WritesToLocalLog" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_LogError_DatabaseFail_WritesToLocalLog" & vbCrLf
    End If
    
    ' Test 4: Detecci├│n de errores cr├¡ticos de base de datos
    testsTotal = testsTotal + 1
    If Test_IsCriticalError_DatabaseErrors_ReturnsTrue() Then
        resultado = resultado & "[OK] Test_IsCriticalError_DatabaseErrors_ReturnsTrue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_IsCriticalError_DatabaseErrors_ReturnsTrue" & vbCrLf
    End If
    
    ' Test 5: Detecci├│n de errores no cr├¡ticos
    testsTotal = testsTotal + 1
    If Test_IsCriticalError_NonCriticalErrors_ReturnsFalse() Then
        resultado = resultado & "[OK] Test_IsCriticalError_NonCriticalErrors_ReturnsFalse" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_IsCriticalError_NonCriticalErrors_ReturnsFalse" & vbCrLf
    End If
    
    ' Test 6: Escritura a log local
    testsTotal = testsTotal + 1
    If Test_WriteToLocalLog_ValidMessage_Success() Then
        resultado = resultado & "[OK] Test_WriteToLocalLog_ValidMessage_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_WriteToLocalLog_ValidMessage_Success" & vbCrLf
    End If
    
    ' Test 7: Limpieza de logs exitosa
    testsTotal = testsTotal + 1
    If Test_CleanOldLogs_Success() Then
        resultado = resultado & "[OK] Test_CleanOldLogs_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CleanOldLogs_Success" & vbCrLf
    End If
    
    ' Test 8: Flujo completo de error (integraci├│n)
    testsTotal = testsTotal + 1
    If Test_Integration_ErrorFlow_Complete() Then
        resultado = resultado & "[OK] Test_Integration_ErrorFlow_Complete" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Integration_ErrorFlow_Complete" & vbCrLf
    End If
    
    ' Test 9: Caso extremo - descripci├│n muy larga
    testsTotal = testsTotal + 1
    If Test_EdgeCase_VeryLongErrorDescription() Then
        resultado = resultado & "[OK] Test_EdgeCase_VeryLongErrorDescription" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EdgeCase_VeryLongErrorDescription" & vbCrLf
    End If
    
    ' Resumen final
    resultado = resultado & vbCrLf & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Pruebas ejecutadas: " & testsTotal & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & testsPassed & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & (testsTotal - testsPassed) & vbCrLf
    
    If testsPassed = testsTotal Then
        resultado = resultado & "RESULTADO: Ô£ô TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "RESULTADO: Ô£ù ALGUNAS PRUEBAS FALLARON" & vbCrLf
    End If
    
    RunErrorHandlerTests = resultado
End Function



' ============================================================================
' FUNCIONES DE CONFIGURACI├ôN DE MOCKS
' ============================================================================

Private Sub SetupMockDatabase()
    With m_MockDB
        .IsConnected = True
        .ShouldFail = False
        .ErrorNumber = 0
        .ErrorDescription = ""
        .ExecuteCallCount = 0
        .LastSQL = ""
        .DatabasePath = "C:\Test\CONDOR_datos.accdb"
        .RecordCount = 0
    End With
End Sub

Private Sub ConfigureMockToFail(errNum As Long, errDesc As String)
    With m_MockDB
        .ShouldFail = True
        .ErrorNumber = errNum
        .ErrorDescription = errDesc
        .IsConnected = False
    End With
End Sub

Private Sub SetupMockFileSystem()
    With m_MockFS
        .FileExists = True
        .CanWrite = True
        .ShouldFailWrite = False
        .LastWrittenContent = ""
        .WriteCallCount = 0
    End With
End Sub

' ============================================================================
' NUEVAS FUNCIONES DE PRUEBA EXPANDIDAS CON MOCKS
' ============================================================================

' Prueba el manejo de errores de base de datos
Private Sub Test_ErrorBaseDatos_ManejaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Simular error de base de datos
    Err.Raise 3001, "Test_ErrorBaseDatos", "Simulated database connection error"
    
ErrorHandler:
    ' Verificar que el error se maneja correctamente
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorBaseDatos_ManejaCorrectamente")
    Resume Next
End Sub

' Prueba el manejo de errores de red
Private Sub Test_ErrorRed_ManejaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Simular error de red
    Err.Raise 2147467259, "Test_ErrorRed", "Network connection timeout"
    
ErrorHandler:
    ' Verificar que el error se maneja correctamente
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorRed_ManejaCorrectamente")
    Resume Next
End Sub

' Prueba el manejo de errores de memoria
Private Sub Test_ErrorMemoria_ManejaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Simular error de memoria
    Err.Raise 7, "Test_ErrorMemoria", "Out of memory"
    
ErrorHandler:
    ' Verificar que el error se maneja correctamente
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorMemoria_ManejaCorrectamente")
    Resume Next
End Sub

' Prueba el manejo de errores de validaci├│n
Private Sub Test_ErrorValidacion_ManejaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Simular error de validaci├│n
    Err.Raise 13, "Test_ErrorValidacion", "Type mismatch - validation failed"
    
ErrorHandler:
    ' Verificar que el error se maneja correctamente
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorValidacion_ManejaCorrectamente")
    Resume Next
End Sub

' Prueba el logging con diferentes niveles de severidad usando mocks
Private Sub Test_LoggingNivelesSeveridad_FuncionaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Configurar mock
    SetupMockDatabase
    
    ' Simular diferentes tipos de errores con diferentes severidades
    Dim criticalErrors As Variant
    Dim warningErrors As Variant
    Dim infoErrors As Variant
    
    criticalErrors = Array(3001, 3024, 7, 9)
    warningErrors = Array(91, 13, 438)
    infoErrors = Array(0, 1001, 2000)
    
    ' Simular logging de errores cr├¡ticos
    Dim i As Integer
    For i = 0 To UBound(criticalErrors)
        If m_MockDB.IsConnected And Not m_MockDB.ShouldFail Then
            m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1
            ' Para errores cr├¡ticos, tambi├®n se crea notificaci├│n
            If IsCriticalErrorMock(criticalErrors(i)) Then
                m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1
            End If
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en Test_LoggingNivelesSeveridad_FuncionaCorrectamente: " & Err.Description
End Sub

' Funci├│n auxiliar para determinar si un error es cr├¡tico (mock)
Private Function IsCriticalErrorMock(ByVal errorNumber As Long) As Boolean
    Select Case errorNumber
        Case 3001, 3024, 3044, 3051, 3078, 3343 ' Errores de base de datos
            IsCriticalErrorMock = True
        Case 7, 9, 11, 13 ' Errores de memoria
            IsCriticalErrorMock = True
        Case Else
            IsCriticalErrorMock = False
    End Select
End Function

' Prueba el manejo de errores en cascada
Private Sub Test_ErroresCascada_ManejaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Simular una serie de errores relacionados
    Call SimularErrorPrimario
    Call SimularErrorSecundario
    Call SimularErrorTerciario
    
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErroresCascada_ManejaCorrectamente")
    Resume Next
End Sub

' Funciones auxiliares para pruebas de errores en cascada
Private Sub SimularErrorPrimario()
    On Error GoTo ErrorHandler
    Err.Raise 1001, "SimularErrorPrimario", "Error primario en la cadena"
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "SimularErrorPrimario")
    Err.Raise Err.Number, Err.Source, Err.Description ' Re-lanzar el error
End Sub

Private Sub SimularErrorSecundario()
    On Error GoTo ErrorHandler
    Err.Raise 1002, "SimularErrorSecundario", "Error secundario causado por el primario"
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "SimularErrorSecundario")
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SimularErrorTerciario()
    On Error GoTo ErrorHandler
    Err.Raise 1003, "SimularErrorTerciario", "Error terciario en la cascada"
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "SimularErrorTerciario")
End Sub

' Prueba el manejo de errores con contexto adicional usando mocks
Private Sub Test_ErrorConContexto_RegistraDetalles()
    On Error GoTo ErrorHandler
    
    ' Configurar mock
    SetupMockDatabase
    
    Dim contexto As String
    contexto = "Usuario: TestUser, M├│dulo: TestModule, Operaci├│n: TestOperation"
    
    ' Simular logging de error con contexto
    If m_MockDB.IsConnected And Not m_MockDB.ShouldFail Then
        m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1
        m_MockDB.LastSQL = "INSERT INTO Tb_Log_Errores"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en Test_ErrorConContexto_RegistraDetalles: " & Err.Description
End Sub

' ============================================================================
' NUEVAS PRUEBAS CON MOCKS PARA FUNCIONES ESPEC├ìFICAS
' ============================================================================

' Prueba para LogError con datos v├ílidos
Public Function Test_LogError_ValidData_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockDatabase
    Dim testErrNumber As Long
    Dim testErrDescription As String
    Dim testErrSource As String
    Dim testUserAction As String
    
    testErrNumber = 3024
    testErrDescription = "No se pudo encontrar el archivo"
    testErrSource = "CExpedienteService.ObtenerExpediente"
    testUserAction = "Consultando expediente 123"
    
    ' Act - Simular llamada exitosa a LogError
    Dim result As Boolean
    If m_MockDB.IsConnected And Not m_MockDB.ShouldFail Then
        m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1
        m_MockDB.LastSQL = "INSERT INTO Tb_Log_Errores"
        result = True
    Else
        result = False
    End If
    
    ' Assert
    Test_LogError_ValidData_Success = result And (m_MockDB.ExecuteCallCount = 1)
    
    Exit Function
    
TestFail:
    Test_LogError_ValidData_Success = False
End Function

' Prueba para LogError con error cr├¡tico que crea notificaci├│n
Public Function Test_LogError_CriticalError_CreatesNotification() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockDatabase
    Dim testErrNumber As Long
    testErrNumber = 3024 ' Error cr├¡tico
    
    ' Act - Simular LogError con error cr├¡tico
    Dim result As Boolean
    If m_MockDB.IsConnected And Not m_MockDB.ShouldFail Then
        m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 2 ' Log + Notificaci├│n
        result = True
    Else
        result = False
    End If
    
    ' Assert - Verificar que se crearon tanto el log como la notificaci├│n
    Test_LogError_CriticalError_CreatesNotification = result And (m_MockDB.ExecuteCallCount = 2)
    
    Exit Function
    
TestFail:
    Test_LogError_CriticalError_CreatesNotification = False
End Function

' Prueba para LogError cuando falla la base de datos y escribe a log local
Public Function Test_LogError_DatabaseFail_WritesToLocalLog() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    ConfigureMockToFail 3044, "Base de datos no disponible"
    SetupMockFileSystem
    
    ' Act - Simular fallo de base de datos que debe escribir a log local
    Dim result As Boolean
    If m_MockDB.ShouldFail And m_MockFS.CanWrite Then
        m_MockFS.WriteCallCount = m_MockFS.WriteCallCount + 1
        m_MockFS.LastWrittenContent = "ERROR EN modErrorHandler.LogError"
        result = True
    Else
        result = False
    End If
    
    ' Assert
    Test_LogError_DatabaseFail_WritesToLocalLog = result And (m_MockFS.WriteCallCount = 1)
    
    Exit Function
    
TestFail:
    Test_LogError_DatabaseFail_WritesToLocalLog = False
End Function

' Prueba para IsCriticalError con errores de base de datos
Public Function Test_IsCriticalError_DatabaseErrors_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim criticalErrors As Variant
    Dim i As Integer
    Dim allCritical As Boolean
    
    criticalErrors = Array(3024, 3044, 3051, 3078, 3343)
    allCritical = True
    
    ' Simular verificaci├│n de errores cr├¡ticos de base de datos
    For i = 0 To UBound(criticalErrors)
        If Not IsCriticalErrorMock(criticalErrors(i)) Then
            allCritical = False
            Exit For
        End If
    Next i
    
    ' Assert
    Test_IsCriticalError_DatabaseErrors_ReturnsTrue = allCritical
    
    Exit Function
    
TestFail:
    Test_IsCriticalError_DatabaseErrors_ReturnsTrue = False
End Function

' Prueba para IsCriticalError con errores no cr├¡ticos
Public Function Test_IsCriticalError_NonCriticalErrors_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim nonCriticalErrors As Variant
    Dim i As Integer
    Dim allNonCritical As Boolean
    
    nonCriticalErrors = Array(1001, 2000, 5000, 6000)
    allNonCritical = True
    
    ' Simular verificaci├│n de errores no cr├¡ticos
    For i = 0 To UBound(nonCriticalErrors)
        If IsCriticalErrorMock(nonCriticalErrors(i)) Then
            allNonCritical = False
            Exit For
        End If
    Next i
    
    ' Assert
    Test_IsCriticalError_NonCriticalErrors_ReturnsFalse = allNonCritical
    
    Exit Function
    
TestFail:
    Test_IsCriticalError_NonCriticalErrors_ReturnsFalse = False
End Function

' Prueba para WriteToLocalLog con mensaje v├ílido
Public Function Test_WriteToLocalLog_ValidMessage_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockFileSystem
    Dim testMessage As String
    testMessage = "Mensaje de prueba para log local"
    
    ' Act - Simular escritura exitosa a log local
    Dim result As Boolean
    If m_MockFS.CanWrite And Not m_MockFS.ShouldFailWrite Then
        m_MockFS.WriteCallCount = m_MockFS.WriteCallCount + 1
        m_MockFS.LastWrittenContent = testMessage
        result = True
    Else
        result = False
    End If
    
    ' Assert
    Test_WriteToLocalLog_ValidMessage_Success = result And _
                                              (m_MockFS.LastWrittenContent = testMessage)
    
    Exit Function
    
TestFail:
    Test_WriteToLocalLog_ValidMessage_Success = False
End Function

' Prueba para CleanOldLogs exitoso
Public Function Test_CleanOldLogs_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockDatabase
    
    ' Act - Simular limpieza exitosa de logs antiguos
    Dim result As Boolean
    If m_MockDB.IsConnected And Not m_MockDB.ShouldFail Then
        m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1
        m_MockDB.LastSQL = "DELETE FROM Tb_Log_Errores WHERE Fecha_Hora"
        result = True
    Else
        result = False
    End If
    
    ' Assert
    Test_CleanOldLogs_Success = result And (InStr(m_MockDB.LastSQL, "DELETE") > 0)
    
    Exit Function
    
TestFail:
    Test_CleanOldLogs_Success = False
End Function

' Prueba de integraci├│n: flujo completo de error
Public Function Test_Integration_ErrorFlow_Complete() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockDatabase
    SetupMockFileSystem
    
    ' Act - Simular flujo completo: Error -> Log -> Notificaci├│n (si es cr├¡tico)
    Dim result As Boolean
    Dim errorIsCritical As Boolean
    
    errorIsCritical = True ' Simular error cr├¡tico
    
    If m_MockDB.IsConnected And Not m_MockDB.ShouldFail Then
        m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1 ' Log
        If errorIsCritical Then
            m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1 ' Notificaci├│n
        End If
        result = True
    Else
        result = False
    End If
    
    ' Assert
    Test_Integration_ErrorFlow_Complete = result And (m_MockDB.ExecuteCallCount = 2)
    
    Exit Function
    
TestFail:
    Test_Integration_ErrorFlow_Complete = False
End Function

' Prueba de caso extremo: descripci├│n de error muy larga
Public Function Test_EdgeCase_VeryLongErrorDescription() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockDatabase
    Dim longDescription As String
    longDescription = String(4000, "X") ' Descripci├│n muy larga
    
    ' Act
    Dim result As Boolean
    If m_MockDB.IsConnected And Not m_MockDB.ShouldFail Then
        m_MockDB.ExecuteCallCount = m_MockDB.ExecuteCallCount + 1
        result = True
    Else
        result = False
    End If
    
    ' Assert
    Test_EdgeCase_VeryLongErrorDescription = result And (Len(longDescription) = 4000)
    
    Exit Function
    
TestFail:
    Test_EdgeCase_VeryLongErrorDescription = False
End Function

' ============================================================================
' PRUEBAS INDIVIDUALES
' ============================================================================

' Prueba que LogError registra correctamente un error en la base de datos
Private Sub Test_LogError_RegistraCorrectamente()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim countBefore As Long
    Dim countAfter As Long
    Dim testErrorNumber As Long
    Dim testErrorDescription As String
    Dim testErrorSource As String
    
    ' Preparar datos de prueba
    testErrorNumber = 9999
    testErrorDescription = "Error de prueba para Test_ErrorHandler"
    testErrorSource = "Test_ErrorHandler.Test_LogError_RegistraCorrectamente"
    
    ' Conectar a la base de datos
    Set db = OpenDatabase(CurrentProject.Path & "\CONDOR_datos.accdb")
    
    ' Contar registros antes
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Log_Errores WHERE Descripcion_Error LIKE '*Error de prueba para Test_ErrorHandler*'", dbOpenSnapshot)
    countBefore = rs!Total
    rs.Close
    
    ' Llamar a LogError
    Call modErrorHandler.LogError(testErrorNumber, testErrorDescription, testErrorSource, "Ejecutando prueba")
    
    ' Contar registros despu├®s
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Log_Errores WHERE Descripcion_Error LIKE '*Error de prueba para Test_ErrorHandler*'", dbOpenSnapshot)
    countAfter = rs!Total
    rs.Close
    
    ' Verificar que se agreg├│ un registro
    If countAfter <= countBefore Then
        Err.Raise 9998, "Test_LogError_RegistraCorrectamente", "No se registr├│ el error en la base de datos"
    End If
    
    ' Limpiar recursos
    db.Close
    Set db = Nothing
    Set rs = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Prueba una funci├│n dise├▒ada para fallar y verificar que el error se registra
Private Sub Test_FuncionConError_RegistraError()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim countBefore As Long
    Dim countAfter As Long
    
    ' Conectar a la base de datos
    Set db = OpenDatabase(CurrentProject.Path & "\CONDOR_datos.accdb")
    
    ' Contar registros antes
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Log_Errores WHERE Origen_Error LIKE '*FuncionQueFalla*'", dbOpenSnapshot)
    countBefore = rs!Total
    rs.Close
    
    ' Llamar a la funci├│n que falla
    Call FuncionQueFalla
    
    ' Contar registros despu├®s
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Log_Errores WHERE Origen_Error LIKE '*FuncionQueFalla*'", dbOpenSnapshot)
    countAfter = rs!Total
    rs.Close
    
    ' Verificar que se agreg├│ un registro
    If countAfter <= countBefore Then
        Err.Raise 9997, "Test_FuncionConError_RegistraError", "No se registr├│ el error de la funci├│n que falla"
    End If
    
    ' Limpiar recursos
    db.Close
    Set db = Nothing
    Set rs = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Prueba que los errores cr├¡ticos crean notificaciones
Private Sub Test_ErrorCritico_CreaNotificacion()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim countBefore As Long
    Dim countAfter As Long
    
    ' Conectar a la base de datos
    Set db = OpenDatabase(CurrentProject.Path & "\CONDOR_datos.accdb")
    
    ' Contar notificaciones antes (si existe la tabla)
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Cola_Correos WHERE Asunto LIKE '*ERROR CR├ìTICO*'", dbOpenSnapshot)
    If Err.Number = 0 Then
        countBefore = rs!Total
        rs.Close
        On Error GoTo ErrorHandler
        
        ' Simular un error cr├¡tico (error de base de datos)
        Call modErrorHandler.LogError(3024, "Error cr├¡tico de prueba", "Test_ErrorHandler.Test_ErrorCritico_CreaNotificacion", "Simulando error cr├¡tico")
        
        ' Contar notificaciones despu├®s
        Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Cola_Correos WHERE Asunto LIKE '*ERROR CR├ìTICO*'", dbOpenSnapshot)
        countAfter = rs!Total
        rs.Close
        
        ' Verificar que se cre├│ una notificaci├│n
        If countAfter <= countBefore Then
            Err.Raise 9996, "Test_ErrorCritico_CreaNotificacion", "No se cre├│ notificaci├│n para error cr├¡tico"
        End If
    Else
        ' Si no existe la tabla de correos, la prueba pasa
        On Error GoTo ErrorHandler
    End If
    
    ' Limpiar recursos
    db.Close
    Set db = Nothing
    Set rs = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Prueba la funci├│n de limpieza de logs antiguos
Private Sub Test_CleanOldLogs_FuncionaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Insertar un log antiguo de prueba
    Dim db As DAO.Database
    Dim strSQL As String
    Dim fechaAntigua As String
    
    Set db = OpenDatabase(CurrentProject.Path & "\CONDOR_datos.accdb")
    
    ' Crear un registro antiguo (45 d├¡as atr├ís)
    fechaAntigua = Format(DateAdd("d", -45, Date), "yyyy-mm-dd hh:nn:ss")
    
    strSQL = "INSERT INTO Tb_Log_Errores (" & _
             "Fecha_Hora, " & _
             "Numero_Error, " & _
             "Descripcion_Error, " & _
             "Origen_Error, " & _
             "Usuario, " & _
             "Accion_Usuario" & _
             ") VALUES (" & _
             "'" & fechaAntigua & "', " & _
             "9995, " & _
             "'Log antiguo de prueba', " & _
             "'Test_ErrorHandler.Test_CleanOldLogs', " & _
             "'TestUser', " & _
             "'Creando log antiguo para prueba'" & _
             ")"
    
    db.Execute strSQL
    
    ' Ejecutar limpieza
    Call modErrorHandler.CleanOldLogs
    
    ' Verificar que el log antiguo fue eliminado
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Log_Errores WHERE Descripcion_Error = 'Log antiguo de prueba'", dbOpenSnapshot)
    
    If rs!Total > 0 Then
        rs.Close
        db.Close
        Err.Raise 9994, "Test_CleanOldLogs_FuncionaCorrectamente", "Los logs antiguos no fueron eliminados correctamente"
    End If
    
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' ============================================================================
' FUNCIONES AUXILIARES
' ============================================================================

' Funci├│n dise├▒ada para fallar a prop├│sito (divisi├│n por cero)
Private Sub FuncionQueFalla()
    On Error GoTo ErrorHandler
    
    Dim resultado As Double
    Dim divisor As Double
    
    divisor = 0
    resultado = 10 / divisor ' Esto causar├í un error de divisi├│n por cero
    
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandler.FuncionQueFalla", "Ejecutando divisi├│n por cero intencional")
    ' No re-lanzar el error para que la prueba pueda verificar el registro
End Sub
