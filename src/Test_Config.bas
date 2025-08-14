Attribute VB_Name = "Test_Config"
' Modulo de Pruebas para Configuracion del Sistema CONDOR
' Pruebas unitarias para el modulo modConfig
' Version: 1.0
' Fecha: 2024

Option Compare Database
Option Explicit

' Funcion principal de pruebas para el modulo de configuracion
Public Function RunAllTests() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DE CONFIGURACION ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Test 1: Inicializacion del entorno
    testsTotal = testsTotal + 1
    If Test_InitializeEnvironment() Then
        resultado = resultado & "[OK] Test_InitializeEnvironment: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FALLO] Test_InitializeEnvironment: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 2: Obtencion de rutas
    testsTotal = testsTotal + 1
    If Test_GetPaths() Then
        resultado = resultado & "[OK] Test_GetPaths: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FALLO] Test_GetPaths: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 3: Modo de desarrollo
    testsTotal = testsTotal + 1
    If Test_DevelopmentMode() Then
        resultado = resultado & "[OK] Test_DevelopmentMode: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FALLO] Test_DevelopmentMode: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 4: Configuracion de estructura
    testsTotal = testsTotal + 1
    If Test_ConfigStructure() Then
        resultado = resultado & "[OK] Test_ConfigStructure: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FALLO] Test_ConfigStructure: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 5: Reset de configuracion
    testsTotal = testsTotal + 1
    If Test_ResetConfiguration() Then
        resultado = resultado & "[OK] Test_ResetConfiguration: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FALLO] Test_ResetConfiguration: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 6: Constante IDAplicacion
    testsTotal = testsTotal + 1
    If Test_IDAplicacionConstant() Then
        resultado = resultado & "[OK] Test_IDAplicacionConstant: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FALLO] Test_IDAplicacionConstant: FALLO" & vbCrLf
        Stop
    End If
    
    ' Test 7: Ruta de base de datos Lanzadera
    testsTotal = testsTotal + 1
    If Test_LanzaderaDbPath() Then
        resultado = resultado & "[OK] Test_LanzaderaDbPath: PASO" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FALLO] Test_LanzaderaDbPath: FALLO" & vbCrLf
        Stop
    End If
    
    ' Resumen final
    resultado = resultado & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Tests ejecutados: " & testsTotal & vbCrLf
    resultado = resultado & "Tests exitosos: " & testsPassed & vbCrLf
    
    If testsPassed = testsTotal Then
        resultado = resultado & "[OK] TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "[FALLO] ALGUNAS PRUEBAS FALLARON" & vbCrLf
    End If
    
    RunAllTests = resultado
End Function

' Test de inicializacion del entorno
Private Function Test_InitializeEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    ' Resetear configuracion antes de probar
    Call ResetConfiguration
    
    ' Inicializar entorno
    Call InitializeEnvironment
    
    ' Verificar que la configuracion se inicializo
    If g_AppConfig.IsInitialized = False Then
        Test_InitializeEnvironment = False
        Exit Function
    End If
    
    ' Verificar que se establecio un entorno activo
    If Len(g_AppConfig.EntornoActivo) = 0 Then
        Test_InitializeEnvironment = False
        Exit Function
    End If
    
    Test_InitializeEnvironment = True
    Exit Function
    
ErrorHandler:
    Test_InitializeEnvironment = False
End Function

' Test de obtencion de rutas
Private Function Test_GetPaths() As Boolean
    On Error GoTo ErrorHandler
    
    ' Asegurar que el entorno este inicializado
    Call InitializeEnvironment
    
    ' Probar GetDataPath
    If Len(GetDataPath()) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    ' Probar GetDatabasePath
    If Len(GetDatabasePath()) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    ' Probar GetExpedientesPath
    If Len(GetExpedientesPath()) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    ' Probar GetPlantillasPath
    If Len(GetPlantillasPath()) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    ' Probar GetLanzaderaDbPath
    If Len(GetLanzaderaDbPath()) = 0 Then
        Test_GetPaths = False
        Exit Function
    End If
    
    Test_GetPaths = True
    Exit Function
    
ErrorHandler:
    Test_GetPaths = False
End Function

' Test del modo de desarrollo
Private Function Test_DevelopmentMode() As Boolean
    On Error GoTo ErrorHandler
    
    ' Asegurar que el entorno este inicializado
    Call InitializeEnvironment
    
    ' El modo de desarrollo debe ser determinable
    Dim isDev As Boolean
    isDev = IsDevelopmentMode()
    
    ' En este contexto, deberia ser modo desarrollo
    If Not isDev Then
        Test_DevelopmentMode = False
        Exit Function
    End If
    
    Test_DevelopmentMode = True
    Exit Function
    
ErrorHandler:
    Test_DevelopmentMode = False
End Function

' Test de estructura de configuracion
Private Function Test_ConfigStructure() As Boolean
    On Error GoTo ErrorHandler
    
    ' Asegurar que el entorno este inicializado
    Call InitializeEnvironment
    
    ' Verificar que la estructura T_AppConfig este correctamente inicializada
    If g_AppConfig.IsInitialized = False Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    ' Verificar que los paths no esten vacios
    If Len(g_AppConfig.DataPath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    If Len(g_AppConfig.DatabasePath) = 0 Then
        Test_ConfigStructure = False
        Exit Function
    End If
    
    Test_ConfigStructure = True
    Exit Function
    
ErrorHandler:
    Test_ConfigStructure = False
End Function

' Test de reset de configuracion
Private Function Test_ResetConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Inicializar primero
    Call InitializeEnvironment
    
    ' Verificar que esta inicializado
    If g_AppConfig.IsInitialized = False Then
        Test_ResetConfiguration = False
        Exit Function
    End If
    
    ' Resetear
    Call ResetConfiguration
    
    ' Verificar que se reseteo
    If g_AppConfig.IsInitialized = True Then
        Test_ResetConfiguration = False
        Exit Function
    End If
    
    Test_ResetConfiguration = True
    Exit Function
    
ErrorHandler:
    Test_ResetConfiguration = False
End Function

' Test de constante IDAplicacion
Private Function Test_IDAplicacionConstant() As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar que la constante este definida y tenga el valor correcto
    If IDAplicacion_CONDOR <> 231 Then
        Test_IDAplicacionConstant = False
        Exit Function
    End If
    
    Test_IDAplicacionConstant = True
    Exit Function
    
ErrorHandler:
    Test_IDAplicacionConstant = False
End Function

' Test de ruta de base de datos Lanzadera
Private Function Test_LanzaderaDbPath() As Boolean
    On Error GoTo ErrorHandler
    
    ' Asegurar que el entorno este inicializado
    Call InitializeEnvironment
    
    ' Obtener la ruta de Lanzadera
    Dim lanzaderaPath As String
    lanzaderaPath = GetLanzaderaDbPath()
    
    ' Verificar que no este vacia
    If Len(lanzaderaPath) = 0 Then
        Test_LanzaderaDbPath = False
        Exit Function
    End If
    
    ' En modo desarrollo, debe contener la ruta local
    If IsDevelopmentMode() Then
        If InStr(lanzaderaPath, "Proyectos\CONDOR") = 0 Then
            Test_LanzaderaDbPath = False
            Exit Function
        End If
    Else
        ' En modo produccion, debe contener la ruta de red
        If InStr(lanzaderaPath, "datoste") = 0 Then
            Test_LanzaderaDbPath = False
            Exit Function
        End If
    End If
    
    Test_LanzaderaDbPath = True
    Exit Function
    
ErrorHandler:
    Test_LanzaderaDbPath = False
End Function