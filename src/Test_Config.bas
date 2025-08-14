' Módulo de Pruebas para Configuración del Sistema CONDOR
' Pruebas unitarias para el módulo modConfig
' Versión: 1.0
' Fecha: 2024

Option Compare Database
Option Explicit

#If DEV_MODE Then

' Función principal de pruebas para el módulo de configuración
Public Function RunAllTests() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DE CONFIGURACION ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Test 1: Inicialización del entorno
    testsTotal = testsTotal + 1
    If Test_InitializeEnvironment() Then
        resultado = resultado & "✓ [OK] Test_InitializeEnvironment: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "✗ [FALLO] Test_InitializeEnvironment: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 2: Obtención de rutas
    testsTotal = testsTotal + 1
    If Test_GetPaths() Then
        resultado = resultado & "✓ [OK] Test_GetPaths: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "✗ [FALLO] Test_GetPaths: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 3: Modo de desarrollo
    testsTotal = testsTotal + 1
    If Test_DevelopmentMode() Then
        resultado = resultado & "✓ [OK] Test_DevelopmentMode: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "✗ [FALLO] Test_DevelopmentMode: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 4: Configuración de estructura
    testsTotal = testsTotal + 1
    If Test_ConfigStructure() Then
        resultado = resultado & "✓ [OK] Test_ConfigStructure: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "✗ [FALLO] Test_ConfigStructure: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 5: Reinicialización
    testsTotal = testsTotal + 1
    If Test_ResetConfiguration() Then
        resultado = resultado & "✓ [OK] Test_ResetConfiguration: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "✗ [FALLO] Test_ResetConfiguration: FALLO" & vbCrLf
        Stop
    End If
    
    ' Resumen final
    resultado = resultado & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Pruebas pasadas: " & testsPassed & "/" & testsTotal & vbCrLf
    
    If testsPassed = testsTotal Then
        resultado = resultado & "✓ TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "✗ ALGUNAS PRUEBAS FALLARON" & vbCrLf
    End If
    
    RunAllTests = resultado
End Function

' Prueba la inicialización del entorno
Private Function Test_InitializeEnvironment() As Boolean
    Dim result As Boolean
    
    ' Ejecutar inicialización
    result = InitializeEnvironment()
    
    ' Verificar que la inicialización fue exitosa
    If Not result Then
        Test_InitializeEnvironment = False
        Exit Function
    End If
    
    ' Verificar que la configuración está marcada como inicializada
    If Not g_AppConfig.IsInitialized Then
        Test_InitializeEnvironment = False
        Exit Function
    End If
    
    Test_InitializeEnvironment = True
End Function

' Prueba la obtención de rutas
Private Function Test_GetPaths() As Boolean
    Dim dbPath As String
    Dim dataPath As String
    Dim sourcePath As String
    Dim backupPath As String
    Dim logPath As String
    Dim tempPath As String
    
    ' Obtener todas las rutas
    dbPath = GetDatabasePath()
    dataPath = GetDataPath()
    sourcePath = GetSourcePath()
    backupPath = GetBackupPath()
    logPath = GetLogPath()
    tempPath = GetTempPath()
    
    ' Verificar que ninguna ruta esté vacía
    If Len(dbPath) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    If Len(dataPath) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    If Len(sourcePath) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    If Len(backupPath) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    If Len(logPath) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    If Len(tempPath) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    ' Verificar que las rutas de desarrollo sean correctas
    If InStr(dbPath, "Proyectos\CONDOR") = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    Test_GetPaths = True
End Function

' Prueba el modo de desarrollo
Private Function Test_DevelopmentMode() As Boolean
    Dim isDev As Boolean
    
    isDev = IsDevelopmentMode()
    
    ' En este contexto, debe estar en modo desarrollo
    If Not isDev Then
        Test_DevelopmentMode = False
        Exit Function
    End If
    
    Test_DevelopmentMode = True
End Function

' Prueba la estructura de configuración
Private Function Test_ConfigStructure() As Boolean
    ' Inicializar si no está inicializado
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    
    ' Verificar que todos los campos de la estructura estén poblados
    If Len(g_AppConfig.DatabasePath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    If Len(g_AppConfig.DataPath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    If Len(g_AppConfig.SourcePath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    If Len(g_AppConfig.BackupPath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    If Len(g_AppConfig.LogPath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    If Len(g_AppConfig.TempPath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    Test_ConfigStructure = True
End Function

' Prueba la reinicialización de configuración
Private Function Test_ResetConfiguration() As Boolean
    ' Primero inicializar
    Call InitializeEnvironment
    
    ' Verificar que está inicializado
    If Not g_AppConfig.IsInitialized Then
        Test_ResetConfiguration = False
        Exit Function
    End If
    
    ' Reinicializar
    Call InitializeEnvironment
    
    ' Verificar que sigue inicializado correctamente
    If Not g_AppConfig.IsInitialized Then
        Test_ResetConfiguration = False
        Exit Function
    End If
    
    ' Verificar que las rutas siguen siendo válidas
    If Len(GetDatabasePath()) = 0 Then
        Test_ResetConfiguration = False
        Exit Function
    End If
    
    Test_ResetConfiguration = True
End Function

#End If