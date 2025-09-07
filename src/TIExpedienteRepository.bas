Attribute VB_Name = "TIExpedienteRepository"
Option Compare Database
Option Explicit

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

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
    On Error GoTo ErrorHandler
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    
    ' Definir nombres de archivos de base de datos
    Dim templateDbName As String: templateDbName = "Expedientes_test_template.accdb"
    Dim activeDbName As String: activeDbName = "Expedientes_integration_test.accdb"
    
    ' Usar la utilidad centralizada para preparar la BD de prueba
    Call modTestUtils.PrepareTestDatabase(templateDbName, activeDbName)
    
    ' Insertar datos de prueba necesarios para la suite
    Dim db As DAO.Database
    Dim activePath As String: activePath = modTestUtils.GetWorkspacePath() & activeDbName
    Dim dbPassword As String: dbPassword = "dpddpd" ' TODO: Obtener desde configuración
    Set db = DBEngine.OpenDatabase(activePath, False, False, ";PWD=" & dbPassword)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo) VALUES (1, 'NEMO-001', 'Expediente de Prueba');", dbFailOnError
    db.Close
    Set db = Nothing
    Exit Sub
ErrorHandler:
    If Not db Is Nothing Then db.Close
    Err.Raise Err.Number, "TIExpedienteRepository.SuiteSetup", Err.Description
End Sub

Private Sub SuiteTeardown()
    ' Limpieza estandarizada a través de la utilidad central.
    Call modTestUtils.CleanupTestDatabase("Expedientes_integration_test.accdb")
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

    ' ARRANGE: Crear una configuración LOCAL apuntando a la BD de prueba de Expedientes
    Dim config As IConfig
    Dim mockConfigImpl As New CMockConfig
    Dim activeDbPath As String: activeDbPath = modTestUtils.GetWorkspacePath() & "Expedientes_integration_test.accdb"
    mockConfigImpl.SetSetting "EXPEDIENTES_DATA_PATH", activeDbPath
    mockConfigImpl.SetSetting "EXPEDIENTES_PASSWORD", "dpddpd"
    Set config = mockConfigImpl

    ' Crear una instancia REAL del repositorio, inyectando la configuración local
    Set repo = modRepositoryFactory.CreateExpedienteRepository(config)
    
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

