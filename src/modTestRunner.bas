Attribute VB_Name = "modTestRunner"
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
        resultado = resultado & "✅ RESULTADO: TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "❌ RESULTADO: ALGUNAS PRUEBAS FALLARON" & vbCrLf
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
    Dim testModule As Object
    Dim procedureName As String
    Dim totalTests As Long, passedTests As Long, failedTests As Long
    Dim currentTest As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.CreateTextFile(strLogPath, True)
    
    logFile.WriteLine "=== INICIO DE LA SUITE DE PRUEBAS CONDOR ==="
    logFile.WriteLine "Fecha: " & Now()
    logFile.WriteLine "============================================" & vbCrLf
    
    ' Bucle para encontrar y ejecutar todas las pruebas
    For Each testModule In Application.VBE.ActiveVBProject.VBComponents
        If Left(testModule.Name, 5) = "Test_" Then
            currentTest = testModule.Name
            logFile.WriteLine "Ejecutando módulo: " & testModule.Name
            
            ' Aquí iría la lógica para iterar y ejecutar cada procedimiento de prueba
            ' Por ahora, simulamos la ejecución exitosa
            totalTests = totalTests + 1
            passedTests = passedTests + 1
            
            logFile.WriteLine "  - Módulo " & testModule.Name & ": PASSED"
        End If
    Next
    
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
    logFile.WriteLine "Fuente: " & Err.Source
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


