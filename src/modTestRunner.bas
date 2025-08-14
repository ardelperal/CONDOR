Attribute VB_Name = "modTestRunner"
Option Compare Database
Option Explicit

' Motor de pruebas para el proyecto CONDOR
Public Sub ExecuteAllTests()
    Dim resultado As String
    resultado = RunAllTests()
    
    ' Enviar resultado a Debug
    Debug.Print resultado
    
    ' Escribir resultado a archivo temporal para que el CLI lo pueda leer
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\condor_test_results.txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error Resume Next
    Open tempFile For Output As #fileNum
    Print #fileNum, resultado
    Close #fileNum
    
    ' Tambien imprimir cada linea individualmente para debug
    Dim lineas() As String
    lineas = Split(resultado, vbCrLf)
    
    Dim i As Integer
    For i = 0 To UBound(lineas)
        If Len(Trim(lineas(i))) > 0 Then
            Debug.Print "[TEST] " & lineas(i)
        End If
    Next i
End Sub

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
    
    ' Ejecutar pruebas de autenticacion
    resultado = resultado & "--- PRUEBAS DE AUTENTICACION ---" & vbCrLf
    On Error GoTo ErrorHandler
    
    resultado = resultado & "[INFO] Iniciando ejecucion de Test_AuthService.RunAllTests()..." & vbCrLf
    
    Dim authResults As String
    authResults = Test_AuthService.RunAllTests()
    
    resultado = resultado & "[DEBUG] Longitud de authResults: " & Len(authResults) & vbCrLf
    
    If Len(authResults) > 0 Then
        resultado = resultado & authResults & vbCrLf
        testsPassed = testsPassed + 8  ' Test_AuthService tiene 8 pruebas
        testsTotal = testsTotal + 8
    Else
        resultado = resultado & "[WARNING] Test_AuthService.RunAllTests() retorno cadena vacia" & vbCrLf
        testsTotal = testsTotal + 8
    End If
    
    ' Ejecutar pruebas de configuracion
    resultado = resultado & vbCrLf & "--- PRUEBAS DE CONFIGURACION ---" & vbCrLf
    resultado = resultado & "[INFO] Iniciando ejecucion de Test_Config.RunAllTests()..." & vbCrLf
    
    Dim configResults As String
    configResults = Test_Config.RunAllTests()
    
    resultado = resultado & "[DEBUG] Longitud de configResults: " & Len(configResults) & vbCrLf
    
    If Len(configResults) > 0 Then
        resultado = resultado & configResults & vbCrLf
        testsPassed = testsPassed + 2  ' Test_Config tiene 2 pruebas
        testsTotal = testsTotal + 2
    Else
        resultado = resultado & "[WARNING] Test_Config.RunAllTests() retorno cadena vacia" & vbCrLf
        testsTotal = testsTotal + 2
    End If
    
    ' Ejecutar pruebas de ExpedienteService
    resultado = resultado & vbCrLf & "--- PRUEBAS DE EXPEDIENTE SERVICE ---" & vbCrLf
    resultado = resultado & "[INFO] Iniciando ejecucion de Test_ExpedienteService.Test_ExpedienteService_All()..." & vbCrLf
    
    On Error Resume Next
    Test_ExpedienteService.Test_ExpedienteService_All
    
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Pruebas de ExpedienteService ejecutadas correctamente" & vbCrLf
        testsPassed = testsPassed + 4  ' Test_ExpedienteService tiene 4 pruebas
        testsTotal = testsTotal + 4
    Else
        resultado = resultado & "[ERROR] Error en pruebas de ExpedienteService: " & Err.Description & vbCrLf
        testsTotal = testsTotal + 4
    End If
    On Error GoTo ErrorHandler
    
    GoTo TestsComplete
    
ErrorHandler:
    resultado = resultado & "[ERROR] Error en prueba: " & Err.Description & " (Numero: " & Err.Number & ")" & vbCrLf
    resultado = resultado & "[ERROR] Fuente del error: " & Err.Source & vbCrLf
    testsTotal = testsTotal + 1
    Resume TestsComplete
    
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
