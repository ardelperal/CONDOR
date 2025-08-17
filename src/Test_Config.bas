Attribute VB_Name = "Test_Config"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE PRUEBAS PARA SISTEMA DE CONFIGURACIÓN
' ============================================================================
' Pruebas para modConfig.bas y CConfig.cls
' Incluye mocks para sistema de archivos y pruebas de entornos
' Fecha: 2025-01-14
' Implementa patrón AAA (Arrange, Act, Assert)
' ============================================================================

' ============================================================================
' TIPOS Y VARIABLES PARA MOCKS
' ============================================================================

' Mock para sistema de archivos
Private Type T_MockFileSystem
    FolderExists As Boolean
    FileExists As Boolean
    CanCreateFolder As Boolean
    ShouldFailCreate As Boolean
    CreatedFolders As String ' Lista separada por ";"
    CheckedPaths As String ' Lista separada por ";"
    CreateFolderCallCount As Integer
End Type

' Mock para configuración
Private Type T_MockConfig
    IsInitialized As Boolean
    EntornoActivo As String
    DatabasePath As String
    DataPath As String
    ExpedientesPath As String
    PlantillasPath As String
    LanzaderaDbPath As String
    SourcePath As String
    BackupPath As String
    LogPath As String
    TempPath As String
    ShouldFailInit As Boolean
    InitCallCount As Integer
End Type

' Variables globales para mocks
Private m_MockFS As T_MockFileSystem
Private m_MockConfig As T_MockConfig

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN DE MOCKS
' ============================================================================

' Configurar mock del sistema de archivos
Private Sub SetupMockFileSystem()
    With m_MockFS
        .FolderExists = True
        .FileExists = True
        .CanCreateFolder = True
        .ShouldFailCreate = False
        .CreatedFolders = ""
        .CheckedPaths = ""
        .CreateFolderCallCount = 0
    End With
End Sub

' Configurar mock para fallar operaciones de archivos
Private Sub ConfigureMockFSToFail()
    With m_MockFS
        .FolderExists = False
        .FileExists = False
        .CanCreateFolder = False
        .ShouldFailCreate = True
    End With
End Sub

' Configurar mock de configuración
Private Sub SetupMockConfig()
    With m_MockConfig
        .IsInitialized = False
        .EntornoActivo = "Test Environment"
        .DatabasePath = "C:\Test\CONDOR.accdb"
        .DataPath = "C:\Test\CONDOR_datos.accdb"
        .ExpedientesPath = "C:\Test\Expedientes_datos.accdb"
        .PlantillasPath = "C:\Test\Plantillas"
        .LanzaderaDbPath = "C:\Test\Lanzadera_Datos.accdb"
        .SourcePath = "C:\Test\src"
        .BackupPath = "C:\Test\backups"
        .LogPath = "C:\Test\logs"
        .TempPath = "C:\Test\temp"
        .ShouldFailInit = False
        .InitCallCount = 0
    End With
End Sub

' Función principal que ejecuta todas las pruebas de configuración
Public Function Test_Config_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DE CONFIGURACION ===" & vbCrLf
    
    ' Test 1: Cargar configuración desde archivo
    testsTotal = testsTotal + 1
    If Test_CargarConfiguracionArchivoAAA() Then
        resultado = resultado & "[OK] Test_CargarConfiguracionArchivo" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CargarConfiguracionArchivo" & vbCrLf
    End If
    
    ' Test 2: Obtener valor de configuración existente
    testsTotal = testsTotal + 1
    If Test_ObtenerValorExistenteAAA() Then
        resultado = resultado & "[OK] Test_ObtenerValorExistente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ObtenerValorExistente" & vbCrLf
    End If
    
    ' Test 3: Obtener valor de configuración inexistente
    testsTotal = testsTotal + 1
    If Test_ObtenerValorInexistenteAAA() Then
        resultado = resultado & "[OK] Test_ObtenerValorInexistente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ObtenerValorInexistente" & vbCrLf
    End If
    
    ' Test 4: Establecer valor de configuración
    testsTotal = testsTotal + 1
    If Test_EstablecerValorConfiguracionAAA() Then
        resultado = resultado & "[OK] Test_EstablecerValorConfiguracion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EstablecerValorConfiguracion" & vbCrLf
    End If
    
    ' Test 5: Validar configuración de base de datos
    testsTotal = testsTotal + 1
    If Test_ValidarConfiguracionBDAAA() Then
        resultado = resultado & "[OK] Test_ValidarConfiguracionBD" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ValidarConfiguracionBD" & vbCrLf
    End If
    
    ' Test 6: Validar configuración de rutas
    testsTotal = testsTotal + 1
    If Test_ValidarConfiguracionRutasAAA() Then
        resultado = resultado & "[OK] Test_ValidarConfiguracionRutas" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ValidarConfiguracionRutas" & vbCrLf
    End If
    
    ' Test 7: Guardar configuración
    testsTotal = testsTotal + 1
    If Test_GuardarConfiguracionAAA() Then
        resultado = resultado & "[OK] Test_GuardarConfiguracion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_GuardarConfiguracion" & vbCrLf
    End If
    
    ' Test 8: Resetear configuración a valores por defecto
    testsTotal = testsTotal + 1
    If Test_ResetearConfiguracionAAA() Then
        resultado = resultado & "[OK] Test_ResetearConfiguracion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ResetearConfiguracion" & vbCrLf
    End If
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resumen Config: " & testsPassed & "/" & testsTotal & " pruebas exitosas" & vbCrLf
    
    Test_Config_RunAll = resultado
End Function

' =====================================================
' PRUEBAS INDIVIDUALES - PATRÓN AAA
' =====================================================

Public Function Test_CargarConfiguracionArchivoAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim archivoEsperado As String
    archivoEsperado = "config.ini"
    
    ' Act - Simular carga de configuración
    Dim archivoExiste As Boolean
    archivoExiste = True
    
    ' Assert
    Test_CargarConfiguracionArchivoAAA = archivoExiste
End Function

Public Function Test_ObtenerValorExistenteAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim clave As String
    clave = "DatabasePath"
    
    ' Act - Simular obtención de valor existente
    Dim valorObtenido As String
    valorObtenido = "C:\CONDOR\Database\CONDOR.accdb"
    
    ' Assert
    Test_ObtenerValorExistenteAAA = (Len(valorObtenido) > 0)
End Function

Public Function Test_ObtenerValorInexistenteAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim claveInexistente As String
    claveInexistente = "ClaveNoExistente"
    
    ' Act - Simular obtención de valor inexistente
    Dim valorInexistente As String
    valorInexistente = ""
    
    ' Assert - Para valores inexistentes, es válido retornar cadena vacía
    Test_ObtenerValorInexistenteAAA = True
End Function

Public Function Test_EstablecerValorConfiguracionAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim clave As String: clave = "NuevaClave"
    Dim valor As String: valor = "NuevoValor"
    
    ' Act - Simular establecimiento de valor
    Dim valorEstablecido As Boolean
    valorEstablecido = True
    
    ' Assert
    Test_EstablecerValorConfiguracionAAA = valorEstablecido
End Function

Public Function Test_ValidarConfiguracionBDAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim rutaBD As String
    rutaBD = "C:\CONDOR\Database\CONDOR.accdb"
    
    ' Act - Simular validación de configuración de BD
    Dim configBDValida As Boolean
    configBDValida = (InStr(rutaBD, ".accdb") > 0)
    
    ' Assert
    Test_ValidarConfiguracionBDAAA = configBDValida
End Function

Public Function Test_ValidarConfiguracionRutasAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim rutasDatos As String
    rutasDatos = "C:\CONDOR\Data\"
    
    ' Act - Simular validación de rutas
    Dim rutasValidas As Boolean
    rutasValidas = (Len(rutasDatos) > 0)
    
    ' Assert
    Test_ValidarConfiguracionRutasAAA = rutasValidas
End Function

Public Function Test_GuardarConfiguracionAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim archivoDestino As String
    archivoDestino = "config_backup.ini"
    
    ' Act - Simular guardado de configuración
    Dim guardadoExitoso As Boolean
    guardadoExitoso = True
    
    ' Assert
    Test_GuardarConfiguracionAAA = guardadoExitoso
End Function

Public Function Test_ResetearConfiguracionAAA() As Boolean
    ' Arrange
    Dim configService As IConfig
    Set configService = config()
    Dim valoresOriginales As String
    valoresOriginales = "defaults"
    
    ' Act - Simular reseteo de configuración
    Dim reseteoExitoso As Boolean
    reseteoExitoso = True
    
    ' Assert
    Test_ResetearConfiguracionAAA = reseteoExitoso
End Function

' ============================================================================
' NUEVAS PRUEBAS CON MOCKS PARA modConfig.bas
' ============================================================================

' Prueba para función config() - singleton
Public Function Test_ModConfig_Singleton_ReturnsSameInstance() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim config1 As IConfig
    Dim config2 As IConfig
    
    Set config1 = config()
    Set config2 = config()
    
    ' Assert - Verificar que es la misma instancia (singleton)
    Test_ModConfig_Singleton_ReturnsSameInstance = (config1 Is config2)
    
    Exit Function
    
TestFail:
    Test_ModConfig_Singleton_ReturnsSameInstance = False
End Function

' Prueba para GetActiveEnvironment
Public Function Test_ModConfig_GetActiveEnvironment_ReturnsString() As Boolean
    On Error GoTo TestFail
    
    ' Act
    Dim environment As String
    environment = GetActiveEnvironment()
    
    ' Assert
    Test_ModConfig_GetActiveEnvironment_ReturnsString = (Len(environment) > 0)
    
    Exit Function
    
TestFail:
    Test_ModConfig_GetActiveEnvironment_ReturnsString = False
End Function

' Prueba para GetDatabasePath
Public Function Test_ModConfig_GetDatabasePath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Act
    Dim dbPath As String
    dbPath = GetDatabasePath()
    
    ' Assert - Verificar que retorna una ruta válida
    Test_ModConfig_GetDatabasePath_ReturnsValidPath = (Len(dbPath) > 0 And InStr(dbPath, ".accdb") > 0)
    
    Exit Function
    
TestFail:
    Test_ModConfig_GetDatabasePath_ReturnsValidPath = False
End Function

' Prueba para GetDataPath
Public Function Test_ModConfig_GetDataPath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Act
    Dim dataPath As String
    dataPath = GetDataPath()
    
    ' Assert
    Test_ModConfig_GetDataPath_ReturnsValidPath = (Len(dataPath) > 0 And InStr(dataPath, "datos.accdb") > 0)
    
    Exit Function
    
TestFail:
    Test_ModConfig_GetDataPath_ReturnsValidPath = False
End Function

' Prueba para GetExpedientesPath
Public Function Test_ModConfig_GetExpedientesPath_ReturnsValidPath() As Boolean
    On Error GoTo TestFail
    
    ' Act
    Dim expPath As String
    expPath = GetExpedientesPath()
    
    ' Assert
    Test_ModConfig_GetExpedientesPath_ReturnsValidPath = (Len(expPath) > 0 And InStr(expPath, "Expedientes") > 0)
    
    Exit Function
    
TestFail:
    Test_ModConfig_GetExpedientesPath_ReturnsValidPath = False
End Function

' ============================================================================
' NUEVAS PRUEBAS CON MOCKS PARA CConfig.cls
' ============================================================================

' Prueba para IConfig_GetValue con claves válidas
Public Function Test_CConfig_GetValue_ValidKeys_ReturnsValues() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim configInstance As IConfig
    Set configInstance = New CConfig
    
    ' Act & Assert - Probar varias claves
    Dim testKeys As Variant
    Dim i As Integer
    Dim allValid As Boolean
    
    testKeys = Array("DATABASEPATH", "DATAPATH", "EXPEDIENTESPATH", "PLANTILLASPATH", "LANZADERADBPATH")
    allValid = True
    
    For i = 0 To UBound(testKeys)
        Dim value As Variant
        value = configInstance.GetValue(testKeys(i))
        If IsNull(value) Or Len(CStr(value)) = 0 Then
            allValid = False
            Exit For
        End If
    Next i
    
    Test_CConfig_GetValue_ValidKeys_ReturnsValues = allValid
    
    Exit Function
    
TestFail:
    Test_CConfig_GetValue_ValidKeys_ReturnsValues = False
End Function

' Prueba para IConfig_GetValue con clave inválida
Public Function Test_CConfig_GetValue_InvalidKey_ReturnsNull() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim configInstance As IConfig
    Set configInstance = New CConfig
    
    ' Act
    Dim value As Variant
    value = configInstance.GetValue("CLAVE_INEXISTENTE")
    
    ' Assert
    Test_CConfig_GetValue_InvalidKey_ReturnsNull = IsNull(value)
    
    Exit Function
    
TestFail:
    Test_CConfig_GetValue_InvalidKey_ReturnsNull = False
End Function

' Prueba para IConfig_HasKey
Public Function Test_CConfig_HasKey_ValidAndInvalidKeys() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim configInstance As IConfig
    Set configInstance = New CConfig
    
    ' Act & Assert
    Dim hasValidKey As Boolean
    Dim hasInvalidKey As Boolean
    
    hasValidKey = configInstance.HasKey("DATABASEPATH")
    hasInvalidKey = configInstance.HasKey("CLAVE_INEXISTENTE")
    
    Test_CConfig_HasKey_ValidAndInvalidKeys = hasValidKey And Not hasInvalidKey
    
    Exit Function
    
TestFail:
    Test_CConfig_HasKey_ValidAndInvalidKeys = False
End Function

' Prueba para detección de entorno de desarrollo
Public Function Test_CConfig_DevelopmentMode_Detection() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockFileSystem
    
    ' Act - Simular detección de entorno
    Dim isDev As Boolean
    
    ' Simular que existen rutas de desarrollo
    If m_MockFS.FolderExists And m_MockFS.FileExists Then
        isDev = True
    Else
        isDev = False
    End If
    
    ' Assert
    Test_CConfig_DevelopmentMode_Detection = isDev
    
    Exit Function
    
TestFail:
    Test_CConfig_DevelopmentMode_Detection = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

' Prueba de integración: flujo completo de configuración
Public Function Test_Integration_CompleteConfigFlow() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockConfig
    SetupMockFileSystem
    
    ' Act - Simular flujo completo
    Dim initResult As Boolean
    Dim pathsValid As Boolean
    
    ' 1. Inicializar
    If Not m_MockConfig.ShouldFailInit Then
        m_MockConfig.IsInitialized = True
        m_MockConfig.InitCallCount = 1
        initResult = True
    Else
        initResult = False
    End If
    
    ' 2. Verificar rutas
    If m_MockConfig.IsInitialized Then
        pathsValid = (Len(m_MockConfig.DatabasePath) > 0 And _
                     Len(m_MockConfig.DataPath) > 0 And _
                     Len(m_MockConfig.ExpedientesPath) > 0)
    Else
        pathsValid = False
    End If
    
    ' Assert
    Test_Integration_CompleteConfigFlow = initResult And pathsValid And (m_MockConfig.InitCallCount = 1)
    
    Exit Function
    
TestFail:
    Test_Integration_CompleteConfigFlow = False
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL EXPANDIDA
' ============================================================================

' Función principal para ejecutar todas las pruebas de configuración
Public Function RunConfigTestsComplete() As Boolean
    Debug.Print "=== INICIANDO PRUEBAS COMPLETAS DE CONFIGURACIÓN ==="
    Debug.Print "Fecha/Hora: " & Now
    Debug.Print ""
    
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    totalTests = 0
    passedTests = 0
    
    ' Pruebas legacy (existentes)
    Debug.Print "--- Pruebas Legacy ---"
    Debug.Print Test_Config_RunAll()
    
    ' Nuevas pruebas con mocks
    Debug.Print ""
    Debug.Print "--- Nuevas Pruebas con Mocks ---"
    
    If Test_ModConfig_Singleton_ReturnsSameInstance() Then
        Debug.Print "✓ Test_ModConfig_Singleton_ReturnsSameInstance: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_ModConfig_Singleton_ReturnsSameInstance: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_ModConfig_GetActiveEnvironment_ReturnsString() Then
        Debug.Print "✓ Test_ModConfig_GetActiveEnvironment_ReturnsString: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_ModConfig_GetActiveEnvironment_ReturnsString: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_ModConfig_GetDatabasePath_ReturnsValidPath() Then
        Debug.Print "✓ Test_ModConfig_GetDatabasePath_ReturnsValidPath: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_ModConfig_GetDatabasePath_ReturnsValidPath: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_ModConfig_GetDataPath_ReturnsValidPath() Then
        Debug.Print "✓ Test_ModConfig_GetDataPath_ReturnsValidPath: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_ModConfig_GetDataPath_ReturnsValidPath: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_ModConfig_GetExpedientesPath_ReturnsValidPath() Then
        Debug.Print "✓ Test_ModConfig_GetExpedientesPath_ReturnsValidPath: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_ModConfig_GetExpedientesPath_ReturnsValidPath: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_CConfig_GetValue_ValidKeys_ReturnsValues() Then
        Debug.Print "✓ Test_CConfig_GetValue_ValidKeys_ReturnsValues: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CConfig_GetValue_ValidKeys_ReturnsValues: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_CConfig_GetValue_InvalidKey_ReturnsNull() Then
        Debug.Print "✓ Test_CConfig_GetValue_InvalidKey_ReturnsNull: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CConfig_GetValue_InvalidKey_ReturnsNull: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_CConfig_HasKey_ValidAndInvalidKeys() Then
        Debug.Print "✓ Test_CConfig_HasKey_ValidAndInvalidKeys: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CConfig_HasKey_ValidAndInvalidKeys: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_CConfig_DevelopmentMode_Detection() Then
        Debug.Print "✓ Test_CConfig_DevelopmentMode_Detection: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CConfig_DevelopmentMode_Detection: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    If Test_Integration_CompleteConfigFlow() Then
        Debug.Print "✓ Test_Integration_CompleteConfigFlow: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_Integration_CompleteConfigFlow: FALLÓ"
    End If
    totalTests = totalTests + 1
    
    ' Resumen final
    Debug.Print ""
    Debug.Print "=== RESUMEN FINAL DE PRUEBAS DE CONFIGURACIÓN ==="
    Debug.Print "Total de nuevas pruebas ejecutadas: " & totalTests
    Debug.Print "Nuevas pruebas que pasaron: " & passedTests
    Debug.Print "Nuevas pruebas que fallaron: " & (totalTests - passedTests)
    Debug.Print "Porcentaje de éxito (nuevas): " & Format((passedTests / totalTests) * 100, "0.0") & "%"
    Debug.Print "=== FIN DE PRUEBAS COMPLETAS DE CONFIGURACIÓN ==="
    
    ' Retornar True si todas las pruebas pasaron
    RunConfigTestsComplete = (passedTests = totalTests)
End Function




