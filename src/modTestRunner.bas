Attribute VB_Name = "modTestRunner"
' Módulo: modTestRunner
' Descripción: Motor principal para ejecutar todas las suites de pruebas del proyecto.

#Const DEV_MODE = True

#If DEV_MODE Then

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
    
    ' --- Ejecutar Pruebas de Configuración ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Configuración ---" & vbCrLf
    suiteResult = Test_Config_RunAll()
    resultado = resultado & suiteResult
    ' (Asumimos que las funciones de prueba devuelven un informe que podemos analizar)
    ' (En una versión futura, podríamos hacer esto más sofisticado)
    
    ' --- Ejecutar Pruebas de Autenticación ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Autenticación ---" & vbCrLf
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
    
    ' --- Ejecutar Pruebas de Integración ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Integración ---" & vbCrLf
    suiteResult = Test_Integracion_RunAll()
    resultado = resultado & suiteResult
    
    ' --- Ejecutar Pruebas de Integración de Solicitudes ---
    resultado = resultado & vbCrLf & "--- Ejecutando Pruebas de Integración de Solicitudes ---" & vbCrLf
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

#End If
