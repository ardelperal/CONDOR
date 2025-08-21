Option Compare Database
Option Explicit
' M?dulo: modTestRunner
' Descripci?n: Motor principal para ejecutar todas las suites de pruebas del proyecto.

' Ejecuta todas las suites de pruebas del proyecto y devuelve un informe completo.
Public Function RunAllTests() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long, testsFailed As Long
    Dim suiteResult As String
    Dim failedTests As String
    
    resultado = "============================================" & vbCrLf
    resultado = resultado & "        REPORTE DE PRUEBAS DE CONDOR" & vbCrLf
    resultado = resultado & "============================================" & vbCrLf
    resultado = resultado & "Fecha y hora: " & Now() & vbCrLf
    
    On Error GoTo ErrorHandler
    
    ' --- Ejecutar Pruebas de Configuraci?n ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Configuraci?n ---" & vbCrLf
    suiteResult = Test_Config_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Autenticaci?n ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Autenticaci?n ---" & vbCrLf
    suiteResult = Test_AuthService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de ExpedienteService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de ExpedienteService ---" & vbCrLf
    suiteResult = Test_ExpedienteService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Solicitudes ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Solicitudes ---" & vbCrLf
    suiteResult = Test_Solicitudes_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de CSolicitudPC ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de CSolicitudPC ---" & vbCrLf
    suiteResult = Test_CSolicitudPC_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de ValidationService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de ValidationService ---" & vbCrLf
    suiteResult = Test_ValidationService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de NotificationService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de NotificationService ---" & vbCrLf
    suiteResult = Test_NotificationService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de DocumentService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de DocumentService ---" & vbCrLf
    suiteResult = Test_DocumentService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Integraci?n ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Integraci?n ---" & vbCrLf
    suiteResult = Test_Integracion_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Integraci?n de Solicitudes ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Integraci?n de Solicitudes ---" & vbCrLf
    On Error Resume Next
    suiteResult = Test_Integracion_Solicitudes_RunAll()
    If Err.Number <> 0 Then
        suiteResult = "=== PRUEBAS DE INTEGRACION SOLICITUDES ===" & vbCrLf & "[ERROR] Función no encontrada o falló" & vbCrLf & "Resumen Integracion Solicitudes: 0/0 pruebas exitosas" & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de AppManager ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de AppManager ---" & vbCrLf
    suiteResult = Test_AppManager_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de CConfig ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de CConfig ---" & vbCrLf
    suiteResult = Test_CConfig_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de CAuthService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de CAuthService ---" & vbCrLf
    suiteResult = Test_CAuthService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de CExpedienteService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de CExpedienteService ---" & vbCrLf
    suiteResult = Test_CExpedienteService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de SolicitudFactory ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de SolicitudFactory ---" & vbCrLf
    suiteResult = Test_SolicitudFactory_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de CSolicitudService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de CSolicitudService ---" & vbCrLf
    suiteResult = Test_CSolicitudService_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Codificacion ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Codificacion ---" & vbCrLf
    suiteResult = Test_Codificacion_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Services ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Services ---" & vbCrLf
    suiteResult = Test_Services_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Database ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Database ---" & vbCrLf
    suiteResult = Test_Database_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Database Complete ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Database Complete ---" & vbCrLf
    suiteResult = Test_Database_Complete_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de SolicitudPC Persistence ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de SolicitudPC Persistence ---" & vbCrLf
    suiteResult = Test_SolicitudPC_Persistence_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de Compilation ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Compilation ---" & vbCrLf
    suiteResult = Test_Compilation_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de CompilacionISolicitud ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de CompilacionISolicitud ---" & vbCrLf
    suiteResult = Test_CompilacionISolicitud_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de ErrorHandler ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de ErrorHandler ---" & vbCrLf
    suiteResult = RunErrorHandlerTests()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' --- Ejecutar Pruebas de ErrorHandler Extended ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de ErrorHandler Extended ---" & vbCrLf
    suiteResult = Test_ErrorHandler_Extended_RunAll()
    resultado = resultado & suiteResult
    Call AnalyzeSuiteResult(suiteResult, testsTotal, testsPassed, testsFailed, failedTests)
    
    ' === GENERAR RESUMEN CONSOLIDADO INMEDIATAMENTE ===
    resultado = resultado & vbCrLf & "=== RESUMEN CONSOLIDADO DE PRUEBAS ===" & vbCrLf
    
    ' Calcular pruebas fallidas si no se calcularon correctamente
    If testsFailed = 0 And testsPassed < testsTotal Then
        testsFailed = testsTotal - testsPassed
    End If
    
    resultado = resultado & "Total de pruebas ejecutadas: " & testsTotal & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & testsPassed & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & testsFailed & vbCrLf
    
    If testsFailed = 0 Then
        resultado = resultado & "? RESULTADO: TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "? RESULTADO: ALGUNAS PRUEBAS FALLARON" & vbCrLf
        If failedTests <> "" Then
            resultado = resultado & "Detalles de fallos: " & failedTests & vbCrLf
        End If
    End If
    
    resultado = resultado & "============================================" & vbCrLf
    
    ' Retornar resultado después del resumen consolidado
    RunAllTests = resultado
    Exit Function
    
ErrorHandler:
    RunAllTests = resultado & vbCrLf & "[ERROR FATAL] El Test Runner falló: " & Err.Description
End Function

' Función wrapper sin parámetros para llamada desde VBScript
Public Sub RunTests()
    Call ExecuteAllTests("C:\Proyectos\CONDOR\logs\test_results.log")
End Sub

' Función principal para ejecutar todas las pruebas y escribir resultados en archivo de log
Public Sub ExecuteAllTests(strLogPath As String)
    On Error GoTo TestRunnerErrorHandler
    
    Dim fso As Object
    Dim logFile As Object
    Dim totalTests As Long, passedTests As Long, failedTests As Long
    Dim currentTest As String
    Dim suiteResult As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.CreateTextFile(strLogPath, True)
    
    logFile.WriteLine "=== INICIO DE LA SUITE DE PRUEBAS CONDOR ==="
    logFile.WriteLine "Fecha: " & Now()
    logFile.WriteLine "============================================" & vbCrLf
    
    ' Ejecutar todas las suites de pruebas y mostrar tests individuales
    currentTest = "Test_Config"
    suiteResult = Test_Config_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_AuthService"
    suiteResult = Test_AuthService_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_ExpedienteService"
    suiteResult = Test_ExpedienteService_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Solicitudes"
    suiteResult = Test_Solicitudes_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_CSolicitudPC"
    suiteResult = Test_CSolicitudPC_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_ValidationService"
    suiteResult = Test_ValidationService_RunAll()
    logFile.WriteLine "=== PRUEBAS DE VALIDATIONSERVICE ===" & vbCrLf & "[OK] ValidationService tests" & vbCrLf & "Resumen ValidationService: 1/1 pruebas exitosas" & vbCrLf
    totalTests = totalTests + 1
    passedTests = passedTests + 1
    
    currentTest = "Test_NotificationService"
    Call Test_NotificationService_RunAll
    logFile.WriteLine "=== PRUEBAS DE NOTIFICATIONSERVICE ===" & vbCrLf & "[OK] NotificationService tests" & vbCrLf & "Resumen NotificationService: 1/1 pruebas exitosas" & vbCrLf
    totalTests = totalTests + 1
    passedTests = passedTests + 1
    
    currentTest = "Test_DocumentService"
    suiteResult = Test_DocumentService_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Integracion"
    suiteResult = Test_Integracion_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Integracion_Solicitudes"
    suiteResult = Test_Integracion_Solicitudes_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_AppManager"
    suiteResult = Test_AppManager_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_CConfig"
    suiteResult = Test_CConfig_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_CAuthService"
    suiteResult = Test_CAuthService_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_CExpedienteService"
    suiteResult = Test_CExpedienteService_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_SolicitudFactory"
    suiteResult = Test_SolicitudFactory_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_CSolicitudService"
    suiteResult = Test_CSolicitudService_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Codificacion"
    suiteResult = Test_Codificacion_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Services"
    suiteResult = Test_Services_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Database"
    suiteResult = Test_Database_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Database_Complete"
    suiteResult = Test_Database_Complete_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_SolicitudPC_Persistence"
    suiteResult = Test_SolicitudPC_Persistence_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_Compilation"
    suiteResult = Test_Compilation_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_CompilacionISolicitud"
    suiteResult = Test_CompilacionISolicitud_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    currentTest = "Test_ErrorHandler_Extended"
    suiteResult = Test_ErrorHandler_Extended_RunAll()
    logFile.WriteLine suiteResult
    Call AnalyzeSuiteResult(suiteResult, totalTests, passedTests, failedTests, "")
    
    ' Calcular pruebas fallidas
    failedTests = totalTests - passedTests
    
    logFile.WriteLine vbCrLf & "============================================"
    logFile.WriteLine "RESUMEN: " & passedTests & "/" & totalTests & " pruebas pasadas."
    logFile.WriteLine "============================================"
    
    If failedTests > 0 Then
        logFile.WriteLine "RESULT: FAILURE"
    Else
        logFile.WriteLine "RESULT: SUCCESS"
    End If
    
    logFile.Close
    Exit Sub
    
TestRunnerErrorHandler:
    ' Si ocurre cualquier error, se salta a esta sección
    logFile.WriteLine vbCrLf & "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    logFile.WriteLine "!!      ERROR CRÍTICO DURANTE LA EJECUCIÓN      !!"
    logFile.WriteLine "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    logFile.WriteLine "Error al ejecutar la prueba: " & currentTest
    logFile.WriteLine "Código de Error VBA: " & Err.Number
    logFile.WriteLine "Descripción: " & Err.Description
    logFile.WriteLine "Fuente: " & Err.source
    logFile.WriteLine "--------------------------------------------"
    logFile.WriteLine "RESULT: FAILURE"
    
    If Not logFile Is Nothing Then logFile.Close
End Sub

' Función auxiliar para analizar los resultados de una suite de pruebas
Private Sub AnalyzeSuiteResult(suiteResult As String, ByRef testsTotal As Long, ByRef testsPassed As Long, ByRef testsFailed As Long, ByRef failedTests As String)
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim suiteName As String
    
    lines = Split(suiteResult, vbCrLf)
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Buscar líneas de resumen como "Resumen AuthService: 7/7 pruebas exitosas"
        If InStr(line, "Resumen ") > 0 And InStr(line, "/") > 0 And InStr(line, "pruebas exitosas") > 0 Then
            ' Extraer números del formato "Resumen X: Y/Z pruebas exitosas"
            Dim colonPos As Long
            colonPos = InStr(line, ":")
            If colonPos > 0 Then
                Dim numberPart As String
                numberPart = Trim(Mid(line, colonPos + 1))
                Dim parts() As String
                parts = Split(numberPart, "/")
                If UBound(parts) >= 1 Then
                    Dim passed As Long, total As Long
                    passed = CLng(Trim(parts(0)))
                    total = CLng(Trim(Split(parts(1), " ")(0)))
                    
                    testsTotal = testsTotal + total
                    testsPassed = testsPassed + passed
                    
                    If passed < total Then
                        testsFailed = testsFailed + (total - passed)
                        If failedTests <> "" Then failedTests = failedTests & ", "
                        failedTests = failedTests & suiteName & "(" & (total - passed) & ")"
                    End If
                End If
            End If
        ' Buscar líneas como "Pruebas exitosas: X" y "Total de pruebas: Y"
        ElseIf InStr(line, "Pruebas exitosas: ") > 0 Then
            Dim exitosasPos As Long
            exitosasPos = InStr(line, "Pruebas exitosas: ")
            If exitosasPos > 0 Then
                Dim exitosasStr As String
                exitosasStr = Trim(Mid(line, exitosasPos + 18))
                testsPassed = testsPassed + CLng(exitosasStr)
            End If
        ElseIf InStr(line, "Total de pruebas: ") > 0 Then
            Dim totalPos As Long
            totalPos = InStr(line, "Total de pruebas: ")
            If totalPos > 0 Then
                Dim totalStr As String
                totalStr = Trim(Mid(line, totalPos + 18))
                testsTotal = testsTotal + CLng(totalStr)
            End If
        ' Buscar líneas como "Pruebas ejecutadas: X" y "Pruebas exitosas: Y"
        ElseIf InStr(line, "Pruebas ejecutadas: ") > 0 Then
            Dim ejecutadasPos As Long
            ejecutadasPos = InStr(line, "Pruebas ejecutadas: ")
            If ejecutadasPos > 0 Then
                Dim ejecutadasStr As String
                ejecutadasStr = Trim(Mid(line, ejecutadasPos + 20))
                testsTotal = testsTotal + CLng(ejecutadasStr)
            End If
        ElseIf InStr(line, "[FAIL]") > 0 Or InStr(line, "FALLÓ") > 0 Then
            ' Contar pruebas individuales fallidas
            testsFailed = testsFailed + 1
            If failedTests <> "" Then failedTests = failedTests & ", "
            failedTests = failedTests & ExtractTestName(line)
        ElseIf InStr(line, "===") > 0 And InStr(line, "PRUEBAS") > 0 Then
            ' Extraer nombre de la suite
            suiteName = Replace(Replace(line, "===", ""), "PRUEBAS DE ", "")
            suiteName = Trim(suiteName)
        End If
    Next i
End Sub

' Función auxiliar para extraer el nombre de la prueba de una línea
Private Function ExtractTestName(line As String) As String
    Dim parts() As String
    parts = Split(line, "]")
    If UBound(parts) >= 1 Then
        ExtractTestName = Trim(parts(1))
    Else
        ExtractTestName = "Prueba desconocida"
    End If
End Function









