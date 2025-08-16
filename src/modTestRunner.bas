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

' Función principal para ejecutar todas las pruebas y escribir resultados en archivo de log
Public Sub ExecuteAllTests(strLogPath As String)
    Dim resultado As String
    Dim fso As Object
    Dim logFile As Object
    Dim testsPassed As Long, testsTotal As Long
    Dim allTestsPassed As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear/abrir archivo de log para escritura
    Set logFile = fso.CreateTextFile(strLogPath, True)
    
    ' Escribir encabezado
    logFile.WriteLine "============================================"
    logFile.WriteLine "        REPORTE DE PRUEBAS DE CONDOR"
    logFile.WriteLine "============================================"
    logFile.WriteLine "Fecha y hora: " & Now()
    logFile.WriteLine ""
    
    allTestsPassed = True
    testsTotal = 0
    testsPassed = 0
    
    ' Ejecutar todas las pruebas de los módulos Test_*
    ' Nota: En una implementación completa, aquí se ejecutarían dinámicamente
    ' todos los procedimientos Public Sub Test_* de todos los módulos Test_*
    
    logFile.WriteLine "--- Ejecutando Pruebas de Compilación ---"
    Call Test_ImplementacionISolicitud
    testsTotal = testsTotal + 1
    testsPassed = testsPassed + 1 ' Asumimos éxito si no hay error
    
    ' Aquí se añadirían más llamadas a las pruebas cuando estén refactorizadas
    ' Call Test_CConfig_Creation_Success
    ' Call Test_CAuthService_Creation_Success
    ' etc.
    
    logFile.WriteLine ""
    logFile.WriteLine "============================================"
    logFile.WriteLine "RESUMEN FINAL:"
    logFile.WriteLine "Pruebas ejecutadas: " & testsTotal
    logFile.WriteLine "Pruebas exitosas: " & testsPassed
    logFile.WriteLine "Pruebas fallidas: " & (testsTotal - testsPassed)
    
    If testsPassed = testsTotal Then
        logFile.WriteLine "RESULT: SUCCESS"
    Else
        logFile.WriteLine "RESULT: FAILURE"
    End If
    
    logFile.WriteLine "============================================"
    
    ' Cerrar archivo
    logFile.Close
    Set logFile = Nothing
    Set fso = Nothing
    
    ' Cerrar Access
    Application.Quit
    
    Exit Sub
    
ErrorHandler:
    If Not logFile Is Nothing Then
        logFile.WriteLine "[ERROR FATAL] El Test Runner falló: " & Err.Description
        logFile.WriteLine "RESULT: FAILURE"
        logFile.Close
    End If
    Set logFile = Nothing
    Set fso = Nothing
    Application.Quit
End Sub


