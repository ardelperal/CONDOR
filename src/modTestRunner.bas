Attribute VB_Name = "modTestRunner"
Option Compare Database
Option Explicit

' M?dulo: modTestRunner
' Descripci?n: Motor principal para ejecutar todas las suites de pruebas del proyecto.

' Ejecuta todas las suites de pruebas del proyecto y devuelve un informe completo.
Public Function RunAllTests() As String
    If Not modConfig.DEV_MODE Then
        RunAllTests = "Las pruebas solo están disponibles en modo desarrollo"
        Exit Function
    End If
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


