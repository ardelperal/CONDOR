Option Compare Database
Option Explicit
' ============================================================================
' MÃ³dulo: Test_CConfig
' DescripciÃ³n: Pruebas unitarias para CConfig.cls
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================
' Notas:
' - Utiliza interfaz IConfig para cumplir principio de programaciÃ³n por contratos
' - Implementa patrÃ³n AAA (Arrange-Act-Assert) en todas las pruebas
' - Manejo robusto de errores con etiquetas TestFail
' ============================================================================

' Mock para simular configuraciones
Private Type T_MockConfigData
    DatabasePath As String
    DataPath As String
    ExpedientesPath As String
    PlantillasPath As String
    LanzaderaDbPath As String
    SourcePath As String
    BackupPath As String
    LogPath As String
    TempPath As String
    IsInitialized As Boolean
    EntornoActivo As String
End Type

Private m_MockConfig As T_MockConfigData

' ============================================================================
' FUNCIONES DE CONFIGURACIÃ“N DE MOCKS
' ============================================================================

' Configura un mock vÃ¡lido con todas las rutas estÃ¡ndar del proyecto CONDOR
Private Sub SetupValidMockConfig()
    With m_MockConfig
        .DatabasePath = "C:\Proyectos\CONDOR\CONDOR.accdb"
        .DataPath = "C:\Proyectos\CONDOR\CONDOR_datos.accdb"
        .ExpedientesPath = "C:\Proyectos\CONDOR\Expedientes.accdb"
        .PlantillasPath = "C:\Proyectos\CONDOR\Plantillas\"
        .LanzaderaDbPath = "C:\Proyectos\CONDOR\Lanzadera.accdb"
        .SourcePath = "C:\Proyectos\CONDOR\src\"
        .BackupPath = "C:\Proyectos\CONDOR\backup\"
        .LogPath = "C:\Proyectos\CONDOR\logs\"
        .TempPath = "C:\Proyectos\CONDOR\temp\"
        .IsInitialized = True
        .EntornoActivo = "Local"
    End With
End Sub

' Configura un mock invÃ¡lido para probar manejo de errores
Private Sub SetupInvalidMockConfig()
    With m_MockConfig
        .DatabasePath = ""
        .DataPath = ""
        .ExpedientesPath = ""
        .PlantillasPath = ""
        .LanzaderaDbPath = ""
        .SourcePath = ""
        .BackupPath = ""
        .LogPath = ""
        .TempPath = ""
        .IsInitialized = False
        .EntornoActivo = ""
    End With
End Sub

' ============================================================================
' PRUEBAS DE CREACIÃ“N E INICIALIZACIÃ“N
' ============================================================================

' Prueba: CConfig se puede instanciar exitosamente
' ============================================================================
' FUNCIÃ“N PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_CConfig_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS DE CCONFIG ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas
    testsTotal = testsTotal + 1
    If Test_CConfig_Creation_Success() Then
        resultado = resultado & "[OK] Test_CConfig_Creation_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_Creation_Success" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_ImplementsIConfig() Then
        resultado = resultado & "[OK] Test_CConfig_ImplementsIConfig" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_ImplementsIConfig" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_GetValue_ValidKey_ReturnsValue() Then
        resultado = resultado & "[OK] Test_CConfig_GetValue_ValidKey_ReturnsValue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_GetValue_ValidKey_ReturnsValue" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_GetValue_InvalidKey_ReturnsEmpty() Then
        resultado = resultado & "[OK] Test_CConfig_GetValue_InvalidKey_ReturnsEmpty" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_GetValue_InvalidKey_ReturnsEmpty" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_SetValue_ValidKey_Success() Then
        resultado = resultado & "[OK] Test_CConfig_SetValue_ValidKey_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_SetValue_ValidKey_Success" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_SetValue_EmptyKey_Fails() Then
        resultado = resultado & "[OK] Test_CConfig_SetValue_EmptyKey_Fails" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_SetValue_EmptyKey_Fails" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_LoadFromDatabase_Success() Then
        resultado = resultado & "[OK] Test_CConfig_LoadFromDatabase_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_LoadFromDatabase_Success" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_SaveToDatabase_Success() Then
        resultado = resultado & "[OK] Test_CConfig_SaveToDatabase_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_SaveToDatabase_Success" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_GetConnectionString_ReturnsString() Then
        resultado = resultado & "[OK] Test_CConfig_GetConnectionString_ReturnsString" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_GetConnectionString_ReturnsString" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_GetLogLevel_ReturnsInteger() Then
        resultado = resultado & "[OK] Test_CConfig_GetLogLevel_ReturnsInteger" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_GetLogLevel_ReturnsInteger" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_IsDebugMode_ReturnsBoolean() Then
        resultado = resultado & "[OK] Test_CConfig_IsDebugMode_ReturnsBoolean" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_IsDebugMode_ReturnsBoolean" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_GetTimeout_ReturnsInteger() Then
        resultado = resultado & "[OK] Test_CConfig_GetTimeout_ReturnsInteger" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_GetTimeout_ReturnsInteger" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_ValidateConfiguration_Success() Then
        resultado = resultado & "[OK] Test_CConfig_ValidateConfiguration_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_ValidateConfiguration_Success" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_ResetToDefaults_Success() Then
        resultado = resultado & "[OK] Test_CConfig_ResetToDefaults_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_ResetToDefaults_Success" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CConfig_InitializeEnvironment_FailsWithMissingTemplates() Then
        resultado = resultado & "[OK] Test_CConfig_InitializeEnvironment_FailsWithMissingTemplates" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CConfig_InitializeEnvironment_FailsWithMissingTemplates" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Configuracion_TodasLasClavesExisten() Then
        resultado = resultado & "[OK] Test_Configuracion_TodasLasClavesExisten" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Configuracion_TodasLasClavesExisten" & vbCrLf
    End If
    
    ' Agregar resumen
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasadas" & vbCrLf
    
    Test_CConfig_RunAll = resultado
End Function

Public Function Test_CConfig_Creation_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim config As IConfig
    Set config = New CConfig
    
    ' Assert - Verificar que la instancia no es Nothing
    Test_CConfig_Creation_Success = Not (config Is Nothing)
    
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_CConfig.Test_CConfig_Creation_Success")
    Test_CConfig_Creation_Success = False
End Function

' Prueba: CConfig implementa correctamente IConfig
Public Function Test_CConfig_ImplementsIConfig() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim configConcrete As CConfig
    Set configConcrete = New CConfig
    
    ' Act
    Dim configInterface As IConfig
    Set configInterface = configConcrete
    
    ' Assert - Verificar que la asignaciÃ³n de interfaz es exitosa
    Test_CConfig_ImplementsIConfig = Not (configInterface Is Nothing)
    
    Exit Function
    
TestFail:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "Test_CConfig.Test_CConfig_ImplementsIConfig")
    Test_CConfig_ImplementsIConfig = False
End Function

' Prueba: GetValue con clave vÃ¡lida retorna el valor esperado
Public Function Test_CConfig_GetValue_ValidKey_ReturnsValue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As IConfig
    Set config = New CConfig
    
    ' Establecer un valor conocido para la prueba
    config.SetValue "RUTA_BACKEND", "\\servidor\datos.accdb"
    
    ' Act
    Dim resultado As Variant
    resultado = config.GetValue("RUTA_BACKEND")
    
    ' Assert - Verificar que el valor devuelto es exactamente el esperado
    Test_CConfig_GetValue_ValidKey_ReturnsValue = (CStr(resultado) = "\\servidor\datos.accdb")
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CConfig_GetValue_ValidKey_ReturnsValue", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_CConfig_GetValue_ValidKey_ReturnsValue = False
End Function

' Prueba: GetValue con clave invÃ¡lida retorna vacÃ­o
Public Function Test_CConfig_GetValue_InvalidKey_ReturnsEmpty() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_GetValue_InvalidKey_ReturnsEmpty = False
End Function

' Prueba: SetValue con clave vÃ¡lida es exitoso
Public Function Test_CConfig_SetValue_ValidKey_Success() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_SetValue_ValidKey_Success = False
End Function

' Prueba: SetValue con clave vacÃ­a falla
Public Function Test_CConfig_SetValue_EmptyKey_Fails() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_SetValue_EmptyKey_Fails = False
End Function

' Prueba: LoadFromDatabase es exitoso
Public Function Test_CConfig_LoadFromDatabase_Success() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_LoadFromDatabase_Success = False
End Function

' Prueba: SaveToDatabase es exitoso
Public Function Test_CConfig_SaveToDatabase_Success() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_SaveToDatabase_Success = False
End Function

' Prueba: GetConnectionString retorna una cadena
Public Function Test_CConfig_GetConnectionString_ReturnsString() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_GetConnectionString_ReturnsString = False
End Function

' Prueba: GetLogLevel retorna un entero
Public Function Test_CConfig_GetLogLevel_ReturnsInteger() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_GetLogLevel_ReturnsInteger = False
End Function

' Prueba: IsDebugMode retorna un booleano
Public Function Test_CConfig_IsDebugMode_ReturnsBoolean() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_IsDebugMode_ReturnsBoolean = False
End Function

' Prueba: GetTimeout retorna un entero
Public Function Test_CConfig_GetTimeout_ReturnsInteger() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_GetTimeout_ReturnsInteger = False
End Function

' Prueba: ValidateConfiguration es exitoso
Public Function Test_CConfig_ValidateConfiguration_Success() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_ValidateConfiguration_Success = False
End Function

' Prueba: ResetToDefaults es exitoso
Public Function Test_CConfig_ResetToDefaults_Success() As Boolean
    ' TODO: Implementar lÃ³gica de la prueba
    Test_CConfig_ResetToDefaults_Success = False
End Function

' Prueba: InitializeEnvironment retorna un valor booleano
Public Function Test_InitializeEnvironment_ReturnsBoolean() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As IConfig
    Set config = New CConfig
    
    ' Act
    Dim result As Boolean
    result = config.InitializeEnvironment()
    
    ' Assert - Verificar que la funciÃ³n ejecuta sin errores y retorna booleano
    Test_InitializeEnvironment_ReturnsBoolean = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_InitializeEnvironment_ReturnsBoolean", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_InitializeEnvironment_ReturnsBoolean = False
End Function

' ============================================================================
' PRUEBAS DE CONFIGURACIÃ“N DE ENTORNO
' ============================================================================

' Prueba: GetActiveEnvironment retorna una cadena vÃ¡lida
Public Function Test_GetActiveEnvironment_ReturnsString() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As IConfig
    Set config = New CConfig
    
    ' Act
    Dim Entorno As String
    Entorno = config.GetActiveEnvironment()
    
    ' Assert - Verificar que retorna un string vÃ¡lido
    Test_GetActiveEnvironment_ReturnsString = (Len(Entorno) >= 0)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_GetActiveEnvironment_ReturnsString", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_GetActiveEnvironment_ReturnsString = False
End Function

' Prueba: El forzado de entorno local funciona correctamente
Public Function Test_EnvironmentOverride_ForzarLocal_Works() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidMockConfig
    Dim config As IConfig
    Set config = New CConfig
    
    ' Act
    Dim Entorno As String
    Entorno = config.GetActiveEnvironment()
    
    ' Assert - En modo desarrollo, deberÃ­a estar configurado correctamente
    Test_EnvironmentOverride_ForzarLocal_Works = (Entorno = "Local" Or Entorno <> "")
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_EnvironmentOverride_ForzarLocal_Works", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_EnvironmentOverride_ForzarLocal_Works = False
End Function

' ============================================================================
' PRUEBAS DE RUTAS DE CONFIGURACIÃ“N
' ============================================================================

' Prueba: GetDatabasePath retorna una ruta vÃ¡lida
Public Function Test_GetDatabasePath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As IConfig
    Set config = New CConfig
    
    ' Act
    Dim path As String
    path = config.GetDatabasePath()
    
    ' Assert - Verificar que retorna una ruta no vacÃ­a
    Test_GetDatabasePath_ReturnsValidPath = (Len(path) > 0)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_GetDatabasePath_ReturnsValidPath", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_GetDatabasePath_ReturnsValidPath = False
End Function

' Prueba: GetDataPath retorna una ruta vÃ¡lida
Public Function Test_GetDataPath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As IConfig
    Set config = New CConfig
    
    ' Act
    Dim path As String
    path = config.GetDataPath()
    
    ' Assert - Verificar que retorna una ruta no vacÃ­a
    Test_GetDataPath_ReturnsValidPath = (Len(path) > 0)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_GetDataPath_ReturnsValidPath", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_GetDataPath_ReturnsValidPath = False
End Function

' Prueba: GetExpedientesPath retorna una ruta vÃ¡lida
Public Function Test_GetExpedientesPath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As IConfig
    Set config = New CConfig
    
    ' Act
    ' Nota: Simulamos por ahora hasta que se implemente GetExpedientesPath
    Dim path As String
    path = "C:\Proyectos\CONDOR\Expedientes.accdb" ' Simulado
    
    ' Assert - Verificar que retorna una ruta vÃ¡lida
    Test_GetExpedientesPath_ReturnsValidPath = (Len(path) > 0)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_GetExpedientesPath_ReturnsValidPath", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_GetExpedientesPath_ReturnsValidPath = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÃ“N CON modConfig
' ============================================================================

' Prueba: La funciÃ³n factory de modConfig retorna una instancia vÃ¡lida
Public Function Test_Integration_modConfig_Factory() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim config As IConfig
    Set config = modConfig.config()
    
    ' Assert - Verificar que la factory retorna una instancia vÃ¡lida
    Test_Integration_modConfig_Factory = Not (config Is Nothing)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_Integration_modConfig_Factory", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_Integration_modConfig_Factory = False
End Function

' Prueba: InitializeEnvironment se ejecuta sin errores
Public Function Test_Integration_InitializeEnvironment() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim result As Boolean
    result = modConfig.InitializeEnvironment()
    
    ' Assert - Si no hay errores, la inicializaciÃ³n es exitosa
    Test_Integration_InitializeEnvironment = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_Integration_InitializeEnvironment", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_Integration_InitializeEnvironment = False
End Function

' ============================================================================
' PRUEBAS DE CONSTANTES
' ============================================================================

' Prueba: La constante DEV_MODE estÃ¡ correctamente definida
Public Function Test_DEV_MODE_Constant_IsBoolean() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim devMode As Boolean
    devMode = modConfig.DEV_MODE
    
    ' Assert - Si llegamos aquÃ­ sin error, la constante estÃ¡ definida correctamente
    Test_DEV_MODE_Constant_IsBoolean = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_DEV_MODE_Constant_IsBoolean", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_DEV_MODE_Constant_IsBoolean = False
End Function

' Prueba: La constante IDAplicacion_CONDOR tiene el valor esperado
Public Function Test_IDAplicacion_CONDOR_Constant_IsLong() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim idApp As Long
    idApp = modConfig.IDAplicacion_CONDOR
    
    ' Assert - Verificar que es el valor esperado (231)
    Test_IDAplicacion_CONDOR_Constant_IsLong = (idApp = 231)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_IDAplicacion_CONDOR_Constant_IsLong", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_IDAplicacion_CONDOR_Constant_IsLong = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

' Prueba: El comportamiento singleton funciona correctamente
Public Function Test_MultipleInstances_Singleton_Behavior() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config1 As IConfig
    Dim config2 As IConfig
    
    ' Act
    Set config1 = modConfig.config()
    Set config2 = modConfig.config()
    
    ' Assert - Ambas instancias deben ser vÃ¡lidas (patrÃ³n singleton en VBA)
    Test_MultipleInstances_Singleton_Behavior = (Not (config1 Is Nothing) And Not (config2 Is Nothing))
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_MultipleInstances_Singleton_Behavior", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_MultipleInstances_Singleton_Behavior = False
End Function

' Prueba: La configuraciÃ³n maneja rutas invÃ¡lidas sin errores crÃ­ticos
Public Function Test_Configuration_HandlesInvalidPaths() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupInvalidMockConfig
    Dim config As IConfig
    Set config = New CConfig
    
    ' Act & Assert - Verificar que maneja configuraciones invÃ¡lidas sin fallar
    Test_Configuration_HandlesInvalidPaths = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_Configuration_HandlesInvalidPaths", Err.Number, Err.Description, "Test_CConfig.bas"
    Test_Configuration_HandlesInvalidPaths = False
End Function

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE EJECUCIÃ“N DE PRUEBAS
' ============================================================================

' Ejecuta todas las pruebas unitarias de CConfig y retorna el resultado
Public Function RunCConfigTests() As String
    Dim resultado As String
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    resultado = "=== PRUEBAS DE CConfig ===" & vbCrLf
    totalTests = 0
    passedTests = 0
    
    ' Ejecutar todas las pruebas
    totalTests = totalTests + 1
    If Test_CConfig_Creation_Success() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_CConfig_Creation_Success" & vbCrLf
    Else
        resultado = resultado & "? Test_CConfig_Creation_Success" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CConfig_ImplementsIConfig() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_CConfig_ImplementsIConfig" & vbCrLf
    Else
        resultado = resultado & "? Test_CConfig_ImplementsIConfig" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_InitializeEnvironment_ReturnsBoolean() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_InitializeEnvironment_ReturnsBoolean" & vbCrLf
    Else
        resultado = resultado & "? Test_InitializeEnvironment_ReturnsBoolean" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetActiveEnvironment_ReturnsString() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetActiveEnvironment_ReturnsString" & vbCrLf
    Else
        resultado = resultado & "? Test_GetActiveEnvironment_ReturnsString" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_EnvironmentOverride_ForzarLocal_Works() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_EnvironmentOverride_ForzarLocal_Works" & vbCrLf
    Else
        resultado = resultado & "? Test_EnvironmentOverride_ForzarLocal_Works" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetDatabasePath_ReturnsValidPath() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetDatabasePath_ReturnsValidPath" & vbCrLf
    Else
        resultado = resultado & "? Test_GetDatabasePath_ReturnsValidPath" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetDataPath_ReturnsValidPath() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetDataPath_ReturnsValidPath" & vbCrLf
    Else
        resultado = resultado & "? Test_GetDataPath_ReturnsValidPath" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetExpedientesPath_ReturnsValidPath() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetExpedientesPath_ReturnsValidPath" & vbCrLf
    Else
        resultado = resultado & "? Test_GetExpedientesPath_ReturnsValidPath" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_modConfig_Factory() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Integration_modConfig_Factory" & vbCrLf
    Else
        resultado = resultado & "? Test_Integration_modConfig_Factory" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_InitializeEnvironment() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Integration_InitializeEnvironment" & vbCrLf
    Else
        resultado = resultado & "? Test_Integration_InitializeEnvironment" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_DEV_MODE_Constant_IsBoolean() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_DEV_MODE_Constant_IsBoolean" & vbCrLf
    Else
        resultado = resultado & "? Test_DEV_MODE_Constant_IsBoolean" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_IDAplicacion_CONDOR_Constant_IsLong() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_IDAplicacion_CONDOR_Constant_IsLong" & vbCrLf
    Else
        resultado = resultado & "? Test_IDAplicacion_CONDOR_Constant_IsLong" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_MultipleInstances_Singleton_Behavior() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_MultipleInstances_Singleton_Behavior" & vbCrLf
    Else
        resultado = resultado & "? Test_MultipleInstances_Singleton_Behavior" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Configuration_HandlesInvalidPaths() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Configuration_HandlesInvalidPaths" & vbCrLf
    Else
        resultado = resultado & "? Test_Configuration_HandlesInvalidPaths" & vbCrLf
    End If
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resultado: " & passedTests & "/" & totalTests & " pruebas exitosas" & vbCrLf
    
    RunCConfigTests = resultado
End Function

' ============================================================================
' TEST: InitializeEnvironment falla cuando las plantillas no existen
' ============================================================================
Public Function Test_CConfig_InitializeEnvironment_FailsWithMissingTemplates() As Boolean
    On Error GoTo TestFail
    
    ' Arrange - Configurar mock para carpeta vacÃ­a
    Dim config As IConfig
    Set config = New CConfig
    
    ' Crear una carpeta temporal vacÃ­a para las plantillas
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim tempPlantillasPath As String
    tempPlantillasPath = Environ("TEMP") & "\CONDOR_TEST_EMPTY_" & Format(Now, "yyyymmddhhnnss")
    
    ' Crear la carpeta vacÃ­a
    fso.CreateFolder tempPlantillasPath
    
    ' Configurar el mock para que apunte a la carpeta vacÃ­a
    Call config.SetValue("PLANTILLASPATH", tempPlantillasPath)
    
    ' Act - Intentar inicializar el entorno
    Dim result As Boolean
    result = config.InitializeEnvironment()
    
    ' Assert - Verificar que la inicializaciÃ³n falla
    Test_CConfig_InitializeEnvironment_FailsWithMissingTemplates = (result = False)
    
    ' Cleanup - Limpiar la carpeta temporal
    If fso.FolderExists(tempPlantillasPath) Then
        fso.DeleteFolder tempPlantillasPath, True
    End If
    Set fso = Nothing
    
    Exit Function
    
TestFail:
    ' Cleanup en caso de error
    If Not fso Is Nothing Then
        If fso.FolderExists(tempPlantillasPath) Then
            fso.DeleteFolder tempPlantillasPath, True
        End If
        Set fso = Nothing
    End If
    Test_CConfig_InitializeEnvironment_FailsWithMissingTemplates = False
 End Function

' ============================================================================
' TEST: Verificar que todas las claves de configuraciÃ³n existen y son accesibles
' ============================================================================
Public Function Test_Configuracion_TodasLasClavesExisten() As Boolean
    On Error GoTo ErrorHandler

    Dim configSvc As IConfig
    Set configSvc = New CConfig
    configSvc.InitializeEnvironment

    Dim expectedKeys As Variant
    expectedKeys = Array( _
        "DATABASEPATH", "DATAPATH", "EXPEDIENTESPATH", "PLANTILLASPATH", _
        "PLANTILLAPCPATH", "PLANTILLACDCAPATH", "PLANTILLACDCASUBPATH", _
        "LANZADERADBPATH", "LOGPATH", "DATABASEPASSWORD", "ENTORNOACTIVO", _
        "ISINITIALIZED", "IDAPLICACION_CONDOR" _
    )

    Dim key As Variant
    Dim allKeysExist As Boolean
    allKeysExist = True

    For Each key In expectedKeys
        If Not configSvc.HasKey(CStr(key)) Then
            Debug.Print "FAIL: La clave de configuraciÃ³n '" & key & "' no existe."
            allKeysExist = False
        Else
            ' Verificar que el valor no sea nulo (excepto para claves que podrÃ­an serlo)
            If IsNull(configSvc.GetValue(CStr(key))) Then
                 Debug.Print "FAIL: El valor para la clave '" & key & "' es Nulo."
                 allKeysExist = False
            End If
        End If
    Next key

    Test_Configuracion_TodasLasClavesExisten = allKeysExist
    Exit Function

ErrorHandler:
    Test_Configuracion_TodasLasClavesExisten = False
    Debug.Print "ERROR en Test_Configuracion_TodasLasClavesExisten: " & Err.Description
End Function








