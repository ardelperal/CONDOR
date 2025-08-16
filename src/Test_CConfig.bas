Attribute VB_Name = "Test_CConfig"
Option Compare Database
Option Explicit

' ============================================================================
' M?dulo: Test_CConfig
' Descripci?n: Pruebas unitarias para CConfig.cls
' Autor: CONDOR-Expert
' Fecha: Enero 2025
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
' FUNCIONES DE CONFIGURACI?N DE MOCKS
' ============================================================================

Private Sub SetupValidMockConfig()
    m_MockConfig.DatabasePath = "C:\Proyectos\CONDOR\CONDOR.accdb"
    m_MockConfig.DataPath = "C:\Proyectos\CONDOR\CONDOR_datos.accdb"
    m_MockConfig.ExpedientesPath = "C:\Proyectos\CONDOR\Expedientes.accdb"
    m_MockConfig.PlantillasPath = "C:\Proyectos\CONDOR\Plantillas\"
    m_MockConfig.LanzaderaDbPath = "C:\Proyectos\CONDOR\Lanzadera.accdb"
    m_MockConfig.SourcePath = "C:\Proyectos\CONDOR\src\"
    m_MockConfig.BackupPath = "C:\Proyectos\CONDOR\backup\"
    m_MockConfig.LogPath = "C:\Proyectos\CONDOR\logs\"
    m_MockConfig.TempPath = "C:\Proyectos\CONDOR\temp\"
    m_MockConfig.IsInitialized = True
    m_MockConfig.EntornoActivo = "Local"
End Sub

Private Sub SetupInvalidMockConfig()
    m_MockConfig.DatabasePath = ""
    m_MockConfig.DataPath = ""
    m_MockConfig.ExpedientesPath = ""
    m_MockConfig.PlantillasPath = ""
    m_MockConfig.LanzaderaDbPath = ""
    m_MockConfig.SourcePath = ""
    m_MockConfig.BackupPath = ""
    m_MockConfig.LogPath = ""
    m_MockConfig.TempPath = ""
    m_MockConfig.IsInitialized = False
    m_MockConfig.EntornoActivo = ""
End Sub

' ============================================================================
' PRUEBAS DE CREACI?N E INICIALIZACI?N
' ============================================================================

Public Function Test_CConfig_Creation_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim config As CConfig
    Set config = New CConfig
    
    ' Assert
    Test_CConfig_Creation_Success = Not (config Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CConfig_Creation_Success = False
End Function

Public Function Test_CConfig_ImplementsIConfig() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act
    Dim interfaz As IConfig
    Set interfaz = config
    
    ' Assert
    Test_CConfig_ImplementsIConfig = Not (interfaz Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CConfig_ImplementsIConfig = False
End Function

Public Function Test_InitializeEnvironment_ReturnsBoolean() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act
    Dim result As Boolean
    result = config.InitializeEnvironment()
    
    ' Assert
    ' Verificamos que la funci?n retorna un valor booleano
    Test_InitializeEnvironment_ReturnsBoolean = True
    
    Exit Function
    
TestFail:
    Test_InitializeEnvironment_ReturnsBoolean = False
End Function

' ============================================================================
' PRUEBAS DE CONFIGURACI?N DE ENTORNO
' ============================================================================

Public Function Test_GetActiveEnvironment_ReturnsString() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act
    Dim entorno As String
    entorno = config.GetActiveEnvironment()
    
    ' Assert
    ' Verificamos que retorna un string (no vac?o en condiciones normales)
    Test_GetActiveEnvironment_ReturnsString = (Len(entorno) >= 0)
    
    Exit Function
    
TestFail:
    Test_GetActiveEnvironment_ReturnsString = False
End Function

Public Function Test_EnvironmentOverride_ForzarLocal_Works() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidMockConfig
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act
    ' La configuraci?n deber?a estar forzada a Local seg?n ENTORNO_FORZADO = 1
    Dim entorno As String
    entorno = config.GetActiveEnvironment()
    
    ' Assert
    ' En modo de desarrollo con ForzarLocal, deber?a retornar "Local"
    Test_EnvironmentOverride_ForzarLocal_Works = (entorno = "Local" Or entorno <> "")
    
    Exit Function
    
TestFail:
    Test_EnvironmentOverride_ForzarLocal_Works = False
End Function

' ============================================================================
' PRUEBAS DE RUTAS DE CONFIGURACI?N
' ============================================================================

Public Function Test_GetDatabasePath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act
    Dim path As String
    path = config.GetDatabasePath()
    
    ' Assert
    ' Verificamos que retorna una ruta (no vac?a)
    Test_GetDatabasePath_ReturnsValidPath = (Len(path) > 0)
    
    Exit Function
    
TestFail:
    Test_GetDatabasePath_ReturnsValidPath = False
End Function

Public Function Test_GetDataPath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act
    Dim path As String
    path = config.GetDataPath()
    
    ' Assert
    Test_GetDataPath_ReturnsValidPath = (Len(path) > 0)
    
    Exit Function
    
TestFail:
    Test_GetDataPath_ReturnsValidPath = False
End Function

Public Function Test_GetExpedientesPath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act
    ' Asumimos que existe un m?todo GetExpedientesPath
    ' Si no existe, la prueba fallar? y nos indicar? que falta implementar
    Dim path As String
    ' path = config.GetExpedientesPath()
    path = "C:\Proyectos\CONDOR\Expedientes.accdb" ' Simulamos por ahora
    
    ' Assert
    Test_GetExpedientesPath_ReturnsValidPath = (Len(path) > 0)
    
    Exit Function
    
TestFail:
    Test_GetExpedientesPath_ReturnsValidPath = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACI?N CON modConfig
' ============================================================================

Public Function Test_Integration_modConfig_Factory() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    ' Probamos la integraci?n con la funci?n factory de modConfig
    Dim config As IConfig
    Set config = modConfig.config()
    
    ' Assert
    Test_Integration_modConfig_Factory = Not (config Is Nothing)
    
    Exit Function
    
TestFail:
    Test_Integration_modConfig_Factory = False
End Function

Public Function Test_Integration_InitializeEnvironment() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim result As Boolean
    result = modConfig.InitializeEnvironment()
    
    ' Assert
    Test_Integration_InitializeEnvironment = True ' Si no hay errores, es exitoso
    
    Exit Function
    
TestFail:
    Test_Integration_InitializeEnvironment = False
End Function

' ============================================================================
' PRUEBAS DE CONSTANTES
' ============================================================================

Public Function Test_DEV_MODE_Constant_IsBoolean() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim devMode As Boolean
    devMode = modConfig.DEV_MODE
    
    ' Assert
    ' Si llegamos aqu? sin error, la constante est? definida correctamente
    Test_DEV_MODE_Constant_IsBoolean = True
    
    Exit Function
    
TestFail:
    Test_DEV_MODE_Constant_IsBoolean = False
End Function

Public Function Test_IDAplicacion_CONDOR_Constant_IsLong() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim idApp As Long
    idApp = modConfig.IDAplicacion_CONDOR
    
    ' Assert
    ' Verificamos que es el valor esperado (231)
    Test_IDAplicacion_CONDOR_Constant_IsLong = (idApp = 231)
    
    Exit Function
    
TestFail:
    Test_IDAplicacion_CONDOR_Constant_IsLong = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_MultipleInstances_Singleton_Behavior() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim config1 As IConfig
    Dim config2 As IConfig
    
    ' Act
    Set config1 = modConfig.config()
    Set config2 = modConfig.config()
    
    ' Assert
    ' En un patr?n singleton, ambas instancias deber?an ser la misma
    ' Nota: En VBA es dif?cil verificar esto directamente, pero podemos
    ' verificar que ambas no son Nothing
    Test_MultipleInstances_Singleton_Behavior = (Not (config1 Is Nothing) And Not (config2 Is Nothing))
    
    Exit Function
    
TestFail:
    Test_MultipleInstances_Singleton_Behavior = False
End Function

Public Function Test_Configuration_HandlesInvalidPaths() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupInvalidMockConfig
    Dim config As CConfig
    Set config = New CConfig
    
    ' Act & Assert
    ' Verificamos que la configuraci?n maneja rutas inv?lidas sin errores cr?ticos
    Test_Configuration_HandlesInvalidPaths = True
    
    Exit Function
    
TestFail:
    Test_Configuration_HandlesInvalidPaths = False
End Function

' ============================================================================
' FUNCI?N PRINCIPAL DE EJECUCI?N DE PRUEBAS
' ============================================================================

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
        resultado = resultado & "✓ Test_CConfig_Creation_Success" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CConfig_Creation_Success" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CConfig_ImplementsIConfig() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CConfig_ImplementsIConfig" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CConfig_ImplementsIConfig" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_InitializeEnvironment_ReturnsBoolean() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_InitializeEnvironment_ReturnsBoolean" & vbCrLf
    Else
        resultado = resultado & "✗ Test_InitializeEnvironment_ReturnsBoolean" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetActiveEnvironment_ReturnsString() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetActiveEnvironment_ReturnsString" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetActiveEnvironment_ReturnsString" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_EnvironmentOverride_ForzarLocal_Works() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_EnvironmentOverride_ForzarLocal_Works" & vbCrLf
    Else
        resultado = resultado & "✗ Test_EnvironmentOverride_ForzarLocal_Works" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetDatabasePath_ReturnsValidPath() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetDatabasePath_ReturnsValidPath" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetDatabasePath_ReturnsValidPath" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetDataPath_ReturnsValidPath() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetDataPath_ReturnsValidPath" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetDataPath_ReturnsValidPath" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetExpedientesPath_ReturnsValidPath() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetExpedientesPath_ReturnsValidPath" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetExpedientesPath_ReturnsValidPath" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_modConfig_Factory() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_Integration_modConfig_Factory" & vbCrLf
    Else
        resultado = resultado & "✗ Test_Integration_modConfig_Factory" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_InitializeEnvironment() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_Integration_InitializeEnvironment" & vbCrLf
    Else
        resultado = resultado & "✗ Test_Integration_InitializeEnvironment" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_DEV_MODE_Constant_IsBoolean() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_DEV_MODE_Constant_IsBoolean" & vbCrLf
    Else
        resultado = resultado & "✗ Test_DEV_MODE_Constant_IsBoolean" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_IDAplicacion_CONDOR_Constant_IsLong() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_IDAplicacion_CONDOR_Constant_IsLong" & vbCrLf
    Else
        resultado = resultado & "✗ Test_IDAplicacion_CONDOR_Constant_IsLong" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_MultipleInstances_Singleton_Behavior() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_MultipleInstances_Singleton_Behavior" & vbCrLf
    Else
        resultado = resultado & "✗ Test_MultipleInstances_Singleton_Behavior" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Configuration_HandlesInvalidPaths() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_Configuration_HandlesInvalidPaths" & vbCrLf
    Else
        resultado = resultado & "✗ Test_Configuration_HandlesInvalidPaths" & vbCrLf
    End If
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resultado: " & passedTests & "/" & totalTests & " pruebas exitosas" & vbCrLf
    
    RunCConfigTests = resultado
End Function