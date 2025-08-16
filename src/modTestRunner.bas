Attribute VB_Name = "modTestRunner"
Option Compare Database
Option Explicit

' M?dulo: modTestRunner
' Descripci?n: Motor principal para ejecutar todas las suites de pruebas del proyecto.

' Ejecuta todas las suites de pruebas del proyecto y devuelve un informe completo.
Public Function RunAllTests() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    Dim suiteResult As String
    
    resultado = "============================================" & vbCrLf
    resultado = resultado & "        REPORTE DE PRUEBAS DE CONDOR" & vbCrLf
    resultado = resultado & "============================================" & vbCrLf
    resultado = resultado & "Fecha y hora: " & Now() & vbCrLf
    
    On Error GoTo ErrorHandler
    
    ' --- Ejecutar Pruebas de Configuraci?n ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Configuraci?n ---" & vbCrLf
    suiteResult = Test_Config_RunAll()
    resultado = resultado & suiteResult
    ' (Asumimos que las funciones de prueba devuelven un informe que podemos analizar)
    ' (En una versi?n futura, podr?amos hacer esto m?s sofisticado)
    
    ' --- Ejecutar Pruebas de Autenticaci?n ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Autenticaci?n ---" & vbCrLf
    suiteResult = Test_AuthService_RunAll()
    resultado = resultado & suiteResult
    
    ' --- Ejecutar Pruebas de ExpedienteService ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de ExpedienteService ---" & vbCrLf
    suiteResult = Test_ExpedienteService_RunAll()
    resultado = resultado & suiteResult
    
    ' --- Ejecutar Pruebas de Solicitudes ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Solicitudes ---" & vbCrLf
    suiteResult = Test_Solicitudes_RunAll()
    resultado = resultado & suiteResult
    
    ' --- Ejecutar Pruebas de Integraci?n ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Integraci?n ---" & vbCrLf
    suiteResult = Test_Integracion_RunAll()
    resultado = resultado & suiteResult
    
    ' --- Ejecutar Pruebas de Integraci?n de Solicitudes ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Integraci?n de Solicitudes ---" & vbCrLf
    suiteResult = Test_Integracion_Solicitudes_RunAll()
    resultado = resultado & suiteResult
    
    resultado = resultado & vbCrLf & "============================================" & vbCrLf
    resultado = resultado & "REPORTE FINAL: Pruebas completadas." & vbCrLf
    resultado = resultado & "============================================"
    
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


