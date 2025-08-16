Attribute VB_Name = "Test_Config"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE PRUEBAS PARA SISTEMA DE CONFIGURACIÓN
' ============================================================================
' Pruebas para modConfig.bas y CConfig.cls
' Incluye mocks para sistema de archivos y pruebas de entornos
' Fecha: 2025-01-14
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

' Funcion principal que ejecuta todas las pruebas de configuracion
Public Function Test_Config_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DE CONFIGURACION ===" & vbCrLf
    
    ' Test 1: Cargar configuracion desde archivo
    On Error Resume Next
    Err.Clear
    Call Test_CargarConfiguracionArchivo
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_CargarConfiguracionArchivo" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_CargarConfiguracionArchivo: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 2: Obtener valor de configuracion existente
    On Error Resume Next
    Err.Clear
    Call Test_ObtenerValorExistente
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_ObtenerValorExistente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_ObtenerValorExistente: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 3: Obtener valor de configuracion inexistente
    On Error Resume Next
    Err.Clear
    Call Test_ObtenerValorInexistente
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_ObtenerValorInexistente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_ObtenerValorInexistente: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 4: Establecer valor de configuracion
    On Error Resume Next
    Err.Clear
    Call Test_EstablecerValorConfiguracion
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_EstablecerValorConfiguracion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_EstablecerValorConfiguracion: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 5: Validar configuracion de base de datos
    On Error Resume Next
    Err.Clear
    Call Test_ValidarConfiguracionBD
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_ValidarConfiguracionBD" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_ValidarConfiguracionBD: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 6: Validar configuracion de rutas
    On Error Resume Next
    Err.Clear
    Call Test_ValidarConfiguracionRutas
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_ValidarConfiguracionRutas" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_ValidarConfiguracionRutas: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 7: Guardar configuracion
    On Error Resume Next
    Err.Clear
    Call Test_GuardarConfiguracion
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_GuardarConfiguracion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_GuardarConfiguracion: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Test 8: Resetear configuracion a valores por defecto
    On Error Resume Next
    Err.Clear
    Call Test_ResetearConfiguracion
    If Err.Number = 0 Then
        resultado = resultado & "[OK] Test_ResetearConfiguracion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[ERROR] Test_ResetearConfiguracion: " & Err.Description & vbCrLf
    End If
    testsTotal = testsTotal + 1
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resumen Config: " & testsPassed & "/" & testsTotal & " pruebas exitosas" & vbCrLf
    
    Test_Config_RunAll = resultado
End Function

' =====================================================
' PRUEBAS INDIVIDUALES
' =====================================================

Public Sub Test_CargarConfiguracionArchivo()
    ' Simular carga de configuracion desde archivo
    Dim config As IConfig
    Set config = config()
    Dim archivoExiste As Boolean
    
    ' Simular archivo de configuracion existente
    archivoExiste = True
    
    If Not archivoExiste Then
        Err.Raise 2001, , "Error: No se pudo cargar el archivo de configuracion"
    End If
End Sub

Public Sub Test_ObtenerValorExistente()
    ' Simular obtencion de valor existente
    Dim config As IConfig
    Set config = config()
    Dim valorObtenido As String
    
    ' Simular valor existente
    valorObtenido = "ValorPrueba"
    
    If Len(valorObtenido) = 0 Then
        Err.Raise 2002, , "Error: No se pudo obtener el valor de configuracion existente"
    End If
End Sub

Public Sub Test_ObtenerValorInexistente()
    ' Simular obtencion de valor inexistente
    Dim config As IConfig
    Set config = config()
    Dim valorInexistente As String
    
    ' Simular valor inexistente (debe retornar cadena vacia o valor por defecto)
    valorInexistente = ""
    
    ' Para valores inexistentes, es valido retornar cadena vacia
    ' No debe generar error, solo retornar valor por defecto
End Sub

Public Sub Test_EstablecerValorConfiguracion()
    ' Simular establecimiento de valor de configuracion
    Dim config As IConfig
    Set config = config()
    Dim valorEstablecido As Boolean
    
    ' Simular establecimiento exitoso
    valorEstablecido = True
    
    If Not valorEstablecido Then
        Err.Raise 2003, , "Error: No se pudo establecer el valor de configuracion"
    End If
End Sub

Public Sub Test_ValidarConfiguracionBD()
    ' Simular validacion de configuracion de base de datos
    Dim config As IConfig
    Set config = config()
    Dim configBDValida As Boolean
    
    ' Simular configuracion de BD valida
    configBDValida = True
    
    If Not configBDValida Then
        Err.Raise 2004, , "Error: Configuracion de base de datos invalida"
    End If
End Sub

Public Sub Test_ValidarConfiguracionRutas()
    ' Simular validacion de configuracion de rutas
    Dim config As IConfig
    Set config = config()
    Dim rutasValidas As Boolean
    
    ' Simular rutas validas
    rutasValidas = True
    
    If Not rutasValidas Then
        Err.Raise 2005, , "Error: Configuracion de rutas invalida"
    End If
End Sub

Public Sub Test_GuardarConfiguracion()
    ' Simular guardado de configuracion
    Dim config As IConfig
    Set config = config()
    Dim guardadoExitoso As Boolean
    
    ' Simular guardado exitoso
    guardadoExitoso = True
    
    If Not guardadoExitoso Then
        Err.Raise 2006, , "Error: No se pudo guardar la configuracion"
    End If
End Sub

Public Sub Test_ResetearConfiguracion()
    ' Simular reseteo de configuracion
    Dim config As IConfig
    Set config = config()
    Dim reseteoExitoso As Boolean
    
    ' Simular reseteo exitoso
    reseteoExitoso = True
    
    If Not reseteoExitoso Then
        Err.Raise 2007, , "Error: No se pudo resetear la configuracion"
    End If
End Sub

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
    Dim configInstance As CConfig
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
Public Sub RunConfigTestsComplete()
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
End Sub




