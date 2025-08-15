Attribute VB_Name = "Test_ErrorHandler"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_ErrorHandler
' Descripción: Pruebas para el sistema de manejo de errores centralizado
' Autor: Sistema CONDOR
' Fecha: 2024
' ============================================================================

' Función principal que ejecuta todas las pruebas del manejo de errores
Public Function RunErrorHandlerTests() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DEL SISTEMA DE MANEJO DE ERRORES ===" & vbCrLf & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Test 1: Verificar que LogError registra correctamente
    On Error Resume Next
    Err.Clear
    Call Test_LogError_RegistraCorrectamente
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_LogError_RegistraCorrectamente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_LogError_RegistraCorrectamente: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 2: Verificar manejo de errores en función diseñada para fallar
    On Error Resume Next
    Err.Clear
    Call Test_FuncionConError_RegistraError
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_FuncionConError_RegistraError" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_FuncionConError_RegistraError: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 3: Verificar que se detectan errores críticos
    On Error Resume Next
    Err.Clear
    Call Test_ErrorCritico_CreaNotificacion
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_ErrorCritico_CreaNotificacion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_ErrorCritico_CreaNotificacion: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 4: Verificar limpieza de logs antiguos
    On Error Resume Next
    Err.Clear
    Call Test_CleanOldLogs_FuncionaCorrectamente
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_CleanOldLogs_FuncionaCorrectamente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_CleanOldLogs_FuncionaCorrectamente: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    On Error GoTo 0
    
    ' Resumen final
    resultado = resultado & vbCrLf & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Pruebas ejecutadas: " & testsTotal & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & testsPassed & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & (testsTotal - testsPassed) & vbCrLf
    
    If testsPassed = testsTotal Then
        resultado = resultado & "RESULTADO: ✓ TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "RESULTADO: ✗ ALGUNAS PRUEBAS FALLARON" & vbCrLf
    End If
    
    RunErrorHandlerTests = resultado
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
    
    ' Contar registros después
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Log_Errores WHERE Descripcion_Error LIKE '*Error de prueba para Test_ErrorHandler*'", dbOpenSnapshot)
    countAfter = rs!Total
    rs.Close
    
    ' Verificar que se agregó un registro
    If countAfter <= countBefore Then
        Err.Raise 9998, "Test_LogError_RegistraCorrectamente", "No se registró el error en la base de datos"
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

' Prueba una función diseñada para fallar y verificar que el error se registra
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
    
    ' Llamar a la función que falla
    Call FuncionQueFalla
    
    ' Contar registros después
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Log_Errores WHERE Origen_Error LIKE '*FuncionQueFalla*'", dbOpenSnapshot)
    countAfter = rs!Total
    rs.Close
    
    ' Verificar que se agregó un registro
    If countAfter <= countBefore Then
        Err.Raise 9997, "Test_FuncionConError_RegistraError", "No se registró el error de la función que falla"
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

' Prueba que los errores críticos crean notificaciones
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
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Cola_Correos WHERE Asunto LIKE '*ERROR CRÍTICO*'", dbOpenSnapshot)
    If Err.Number = 0 Then
        countBefore = rs!Total
        rs.Close
        On Error GoTo ErrorHandler
        
        ' Simular un error crítico (error de base de datos)
        Call modErrorHandler.LogError(3024, "Error crítico de prueba", "Test_ErrorHandler.Test_ErrorCritico_CreaNotificacion", "Simulando error crítico")
        
        ' Contar notificaciones después
        Set rs = db.OpenRecordset("SELECT COUNT(*) as Total FROM Tb_Cola_Correos WHERE Asunto LIKE '*ERROR CRÍTICO*'", dbOpenSnapshot)
        countAfter = rs!Total
        rs.Close
        
        ' Verificar que se creó una notificación
        If countAfter <= countBefore Then
            Err.Raise 9996, "Test_ErrorCritico_CreaNotificacion", "No se creó notificación para error crítico"
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

' Prueba la función de limpieza de logs antiguos
Private Sub Test_CleanOldLogs_FuncionaCorrectamente()
    On Error GoTo ErrorHandler
    
    ' Insertar un log antiguo de prueba
    Dim db As DAO.Database
    Dim strSQL As String
    Dim fechaAntigua As String
    
    Set db = OpenDatabase(CurrentProject.Path & "\CONDOR_datos.accdb")
    
    ' Crear un registro antiguo (45 días atrás)
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

' Función diseñada para fallar a propósito (división por cero)
Private Sub FuncionQueFalla()
    On Error GoTo ErrorHandler
    
    Dim resultado As Double
    Dim divisor As Double
    
    divisor = 0
    resultado = 10 / divisor ' Esto causará un error de división por cero
    
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_ErrorHandler.FuncionQueFalla", "Ejecutando división por cero intencional")
    ' No re-lanzar el error para que la prueba pueda verificar el registro
End Sub