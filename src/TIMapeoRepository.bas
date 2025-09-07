Attribute VB_Name = "TIMapeoRepository"
Option Compare Database
Option Explicit

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TIMapeoRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIMapeoRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestGetMapeoPorTipoSuccess()
    suiteResult.AddResult TestGetMapeoPorTipoNotFound()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIMapeoRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler
    ' Aprovisionar la base de datos para toda la suite
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    
    ' Usar las constantes ya definidas para construir los nombres de archivo
    Dim templateDbName As String: templateDbName = "Mapeo_test_template.accdb"
    Dim activeDbName As String: activeDbName = "Mapeo_integration_test.accdb"
    
    ' Llamada al método correcto de modTestUtils
    modTestUtils.PrepareTestDatabase templateDbName, activeDbName
    
    ' Insertar los datos de prueba maestros para la suite
    Dim db As DAO.Database
    Dim activePath As String: activePath = modTestUtils.GetWorkspacePath() & activeDbName
    Set db = DBEngine.OpenDatabase(activePath)
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'refContrato', 'MARCADOR_CONTRATO')", dbFailOnError
    db.Close
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TIMapeoRepository.SuiteSetup", Err.Description
End Sub

Private Sub SuiteTeardown()
    ' Limpieza estandarizada a través de la utilidad central.
    Call modTestUtils.CleanupTestDatabase("Mapeo_integration_test.accdb")
End Sub

' ============================================================================
' TESTS INDIVIDUALES (SE AÑADIRÁN EN LOS SIGUIENTES PROMPTS)
' ============================================================================

Private Function TestGetMapeoPorTipoSuccess() As CTestResult
    Set TestGetMapeoPorTipoSuccess = New CTestResult
    TestGetMapeoPorTipoSuccess.Initialize "GetMapeoPorTipo debe devolver un objeto EMapeo con datos"
    
    Dim localConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim repository As IMapeoRepository
    Dim mapeoResult As EMapeo
    Dim db As DAO.Database
    Dim fs As IFileSystem
    Dim dbPath As String
    
    On Error GoTo TestFail

    ' Arrange: Crear configuración local apuntando a la BD de prueba de esta suite
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "CONDOR_DATA_PATH", modTestUtils.GetWorkspacePath() & "Mapeo_integration_test.accdb"
    mockConfigImpl.SetSetting "CONDOR_PASSWORD", ""
    Set localConfig = mockConfigImpl
    
    ' Arrange: Crear dependencias inyectando la configuración local
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set repository = modRepositoryFactory.CreateMapeoRepository(localConfig)
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = modTestUtils.GetWorkspacePath() & "Mapeo_integration_test.accdb"
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 100, "Test.Arrange", "La BD de prueba de Mapeo no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath, False, False)

    ' Act
    DBEngine.BeginTrans
    Set mapeoResult = repository.GetMapeoPorTipo("PC") ' Este dato se insertó en SuiteSetup
    
    ' Assert
    modAssert.AssertNotNull mapeoResult, "El objeto EMapeo no debe ser nulo."
    modAssert.AssertEquals "PC", mapeoResult.NombrePlantilla, "El nombre de la plantilla no es el esperado."

    TestGetMapeoPorTipoSuccess.Pass
    GoTo Cleanup

TestFail:
    TestGetMapeoPorTipoSuccess.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set mapeoResult = Nothing
    Set repository = Nothing
    Set errorHandler = Nothing
    Set localConfig = Nothing
    Set db = Nothing
    Set fs = Nothing
End Function

Private Function TestGetMapeoPorTipoNotFound() As CTestResult
    Set TestGetMapeoPorTipoNotFound = New CTestResult
    TestGetMapeoPorTipoNotFound.Initialize "GetMapeoPorTipo debe devolver Nothing si no hay mapeo"
    
    Dim localConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim repository As IMapeoRepository
    Dim mapeoResult As EMapeo
    Dim db As DAO.Database
    Dim fs As IFileSystem
    Dim dbPath As String
    
    On Error GoTo TestFail

    ' Arrange: Crear configuración local apuntando a la BD de prueba de esta suite
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "CONDOR_DATA_PATH", modTestUtils.GetWorkspacePath() & "Mapeo_integration_test.accdb"
    mockConfigImpl.SetSetting "CONDOR_PASSWORD", ""
    Set localConfig = mockConfigImpl
    
    ' Arrange: Crear dependencias inyectando la configuración local
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set repository = modRepositoryFactory.CreateMapeoRepository(localConfig)
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = modTestUtils.GetWorkspacePath() & "Mapeo_integration_test.accdb"
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 100, "Test.Arrange", "La BD de prueba de Mapeo no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath, False, False)

    ' Act
    DBEngine.BeginTrans
    Set mapeoResult = repository.GetMapeoPorTipo("TIPO_INEXISTENTE")
    
    ' Assert
    modAssert.AssertIsNull mapeoResult, "El objeto EMapeo devuelto debería ser Nothing."

    TestGetMapeoPorTipoNotFound.Pass
    GoTo Cleanup

TestFail:
    TestGetMapeoPorTipoNotFound.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set mapeoResult = Nothing
    Set repository = Nothing
    Set errorHandler = Nothing
    Set localConfig = Nothing
    Set db = Nothing
    Set fs = Nothing
End Function

