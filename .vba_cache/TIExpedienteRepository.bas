Attribute VB_Name = "TIExpedienteRepository"
Option Compare Database
Option Explicit

' --- Constantes del entorno de prueba ---
Private Const DB_TEMPLATE_NAME As String = "Expedientes_test_template.accdb"
Private Const DB_ACTIVE_NAME As String = "Expedientes_integration_test.accdb"

' =====================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' =====================================================
Public Function TIExpedienteRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIExpedienteRepository - Pruebas de Integración (Optimizadas)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult Test_ObtenerExpedientePorId_Exitoso()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    Set TIExpedienteRepositoryRunAll = suiteResult
End Function

' =====================================================
' GESTIÓN DEL ENTORNO DE LA SUITE
' =====================================================
Private Sub SuiteSetup()
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & "back\test_db\templates\" & DB_TEMPLATE_NAME
    Dim activePath As String: activePath = projectPath & "back\test_db\active\" & DB_ACTIVE_NAME
    
    ' Usar la utilidad centralizada para preparar la BD de prueba
    Call modTestUtils.PrepareTestDatabase(templatePath, activePath)
    
    ' Insertar datos de prueba necesarios para la suite
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(activePath, False, False, ";PWD=dpddpd")
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo) VALUES (1, 'NEMO-001', 'Expediente de Prueba');", dbFailOnError
    db.Close
    Set db = Nothing
End Sub

Private Sub SuiteTeardown()
    On Error Resume Next
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & "back\test_db\active\" & DB_ACTIVE_NAME
    Dim fs As IFileSystem: Set fs = modFileSystemFactory.CreateFileSystem()
    If fs.FileExists(activePath) Then fs.DeleteFile activePath
    Set fs = Nothing
End Sub

' =====================================================
' TESTS INDIVIDUALES
' =====================================================
Private Function Test_ObtenerExpedientePorId_Exitoso() As CTestResult
    Set Test_ObtenerExpedientePorId_Exitoso = New CTestResult
    Test_ObtenerExpedientePorId_Exitoso.Initialize "ObtenerExpedientePorId con un ID existente debe devolver un objeto EExpediente poblado"

    Dim repo As IExpedienteRepository
    Dim result As EExpediente
    
    On Error GoTo TestFail

    ' ARRANGE: El repositorio obtiene la configuración del contexto centralizado
    Set repo = modRepositoryFactory.CreateExpedienteRepository()
    
    ' ACT
    Set result = repo.ObtenerExpedientePorId(1) ' El ID 1 fue insertado en SuiteSetup
    
    ' ASSERT
    modAssert.AssertNotNull result, "El expediente devuelto no debe ser Nulo."
    modAssert.AssertEquals "NEMO-001", result.Nemotecnico, "El nemotécnico del expediente debe ser el correcto."
    
    Test_ObtenerExpedientePorId_Exitoso.Pass
    GoTo Cleanup

TestFail:
    Test_ObtenerExpedientePorId_Exitoso.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set repo = Nothing
    Set result = Nothing
End Function

