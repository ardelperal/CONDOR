Attribute VB_Name "modTestRunner"
Option Compare Database
Option Explicit

' Motor de pruebas para el proyecto CONDOR
Public Function RunAllTests() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    ' Deshabilitar avisos de Access
    Application.Echo False
    DoCmd.SetWarnings False
    
    resultado = "============================================" & vbCrLf
    resultado = resultado & "        PRUEBAS DE CONDOR" & vbCrLf
    resultado = resultado & "============================================" & vbCrLf & vbCrLf
    
    resultado = resultado & "Fecha y hora: " & Now() & vbCrLf
    resultado = resultado & "Base de datos: " & CurrentDb.Name & vbCrLf & vbCrLf
    
    ' Ejecutar pruebas de configuracion
    resultado = resultado & "--- PRUEBAS DE CONFIGURACION ---" & vbCrLf
    On Error GoTo ErrorHandler
    
    Call Test_Config.Test_ConfigModule
    resultado = resultado & "[OK] Test_ConfigModule: PASO" & vbCrLf
    testsPassed = testsPassed + 1
    testsTotal = testsTotal + 1
    
    Call Test_Config.Test_DevelopmentPaths
    resultado = resultado & "[OK] Test_DevelopmentPaths: PASO" & vbCrLf
    testsPassed = testsPassed + 1
    testsTotal = testsTotal + 1
    
    Call Test_Config.Test_UtilityFunctions
    resultado = resultado & "[OK] Test_UtilityFunctions: PASO" & vbCrLf
    testsPassed = testsPassed + 1
    testsTotal = testsTotal + 1
    
    Call Test_Config.Test_AutoInitialization
    resultado = resultado & "[OK] Test_AutoInitialization: PASO" & vbCrLf
    testsPassed = testsPassed + 1
    testsTotal = testsTotal + 1
    
    Call Test_Config.Test_PathConsistency
    resultado = resultado & "[OK] Test_PathConsistency: PASO" & vbCrLf
    testsPassed = testsPassed + 1
    testsTotal = testsTotal + 1
    
    GoTo TestsComplete
    
ErrorHandler:
    resultado = resultado & "[ERROR] Error en prueba: " & Err.Description & vbCrLf
    testsTotal = testsTotal + 1
    Resume Next
    
TestsComplete:
    resultado = resultado & vbCrLf & "RESUMEN:" & vbCrLf
    resultado = resultado & "Pruebas ejecutadas: " & testsTotal & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & testsPassed & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & (testsTotal - testsPassed) & vbCrLf & vbCrLf
    
    If testsPassed = testsTotal Then
        resultado = resultado & "RESULTADO: TODAS LAS PRUEBAS PASARON [OK]" & vbCrLf
    Else
        resultado = resultado & "RESULTADO: ALGUNAS PRUEBAS FALLARON [ERROR]" & vbCrLf
    End If
    
    resultado = resultado & "============================================" & vbCrLf
    
    ' Restaurar configuracion
    Application.Echo True
    DoCmd.SetWarnings True
    
    RunAllTests = resultado
End Function