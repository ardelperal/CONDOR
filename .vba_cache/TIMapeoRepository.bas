Attribute VB_Name = "TIMapeoRepository"
Option Compare Database
Option Explicit

Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\active\CONDOR_mapeo_integration_test.accdb"

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
    Dim templatePath As String: templatePath = projectPath & CONDOR_TEMPLATE_PATH
    Dim activePath As String: activePath = projectPath & CONDOR_ACTIVE_PATH
    Call modTestUtils.SuiteSetup(templatePath, activePath)
    
    ' Insertar los datos de prueba maestros para la suite
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(activePath)
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'refContrato', 'MARCADOR_CONTRATO')", dbFailOnError
    db.Close
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TIMapeoRepository.SuiteSetup", Err.Description
End Sub

Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Call modTestUtils.SuiteTeardown(activePath)
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

    ' Arrange: 1. Crear configuración local apuntando a la BD de prueba activa
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Set localConfig = mockConfigImpl
    
    ' Arrange: 2. Crear dependencias inyectando la configuración local
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(localConfig)
    Set repository = modRepositoryFactory.CreateMapeoRepository(localConfig, errorHandler)
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = localConfig.GetDataPath()
    
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

    ' Arrange
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Set localConfig = mockConfigImpl
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(localConfig)
    Set repository = modRepositoryFactory.CreateMapeoRepository(localConfig, errorHandler)
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = localConfig.GetDataPath()
    
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

