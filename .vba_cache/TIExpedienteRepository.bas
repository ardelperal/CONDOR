Attribute VB_Name = "TIExpedienteRepository"
Option Compare Database
Option Explicit


Private Const EXPEDIENTES_DB_TEMPLATE_PATH As String = "back\test_db\templates\Expedientes_test_template.accdb"
Private Const EXPEDIENTES_DB_ACTIVE_PATH As String = "back\test_db\active\Expedientes_itest.accdb"

Public Function TIExpedienteRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIExpedienteRepository - Pruebas de Integración (Optimizadas)"
    
    On Error GoTo CleanupSuite

    ' 1. Configurar el entorno UNA SOLA VEZ para toda la suite
    Call SuiteSetup
    
    ' 2. Ejecutar todas las pruebas individuales (sin setup/teardown)
    suiteResult.AddResult Test_ObtenerExpedientePorId_Exitoso()
    suiteResult.AddResult Test_ObtenerExpedientePorId_NoEncontradoDevuelveNothing()
    
    ' 3. Limpiar el entorno UNA SOLA VEZ al final
CleanupSuite:
    Call SuiteTeardown
    
    ' Si ocurrió un error durante el Setup o la ejecución, registrarlo
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIExpedienteRepositoryRunAll = suiteResult
End Function

Private Sub SuiteSetup()
    ' Llama a la utilidad centralizada para crear la BD de prueba para toda la suite
    Dim projectPath As String
    projectPath = modTestUtils.GetProjectPath()
    
    Dim templatePath As String
    templatePath = projectPath & EXPEDIENTES_DB_TEMPLATE_PATH
    
    Dim activePath As String
    activePath = projectPath & EXPEDIENTES_DB_ACTIVE_PATH
    
    Call modTestUtils.SuiteSetup(templatePath, activePath)
End Sub

Private Sub SuiteTeardown()
    ' Llama a la utilidad centralizada para eliminar la BD de prueba de la suite
    Dim activePath As String
    activePath = modTestUtils.GetProjectPath() & EXPEDIENTES_DB_ACTIVE_PATH
    
    Call modTestUtils.SuiteTeardown(activePath)
End Sub

Private Function Test_ObtenerExpedientePorId_Exitoso() As CTestResult
    Set Test_ObtenerExpedientePorId_Exitoso = New CTestResult
    Test_ObtenerExpedientePorId_Exitoso.Initialize "ObtenerExpedientePorId con un ID existente debe devolver un objeto EExpediente poblado"
    
    Dim config As IConfig
    Dim repo As IExpedienteRepository
    Dim result As EExpediente
    Dim db As DAO.Database
    Dim fs As IFileSystem
    
    On Error GoTo TestFail

    ' Arrange
    ' 1. Obtener la configuración de prueba (que apunta a la BD activa)
    Set config = modTestContext.GetTestConfig()
    
    ' 2. Crear una instancia REAL del repositorio a probar
    Set repo = modRepositoryFactory.CreateExpedienteRepository(config)
    
    ' 3. Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim dbPath As String: dbPath = config.GetDataPath()
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 100, "Test.Arrange", "La BD de prueba de Expedientes no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath, False, False)
    
    ' Act
    ' 4. Iniciar una transacción para aislar la prueba
    DBEngine.BeginTrans
    
    ' 5. Ejecutar el método bajo prueba
    Set result = repo.ObtenerExpedientePorId(1) ' Asumimos que el ID 1 existe en la plantilla

    ' Assert
    ' 6. Realizar las aserciones
    modAssert.AssertNotNull result, "El expediente devuelto no debe ser Nulo."
    modAssert.AssertEquals 1, result.idExpediente, "El ID del expediente no coincide."
    modAssert.AssertEquals "EXP2023-001", result.Nemotecnico, "El nemotécnico del expediente no coincide."
    
    Test_ObtenerExpedientePorId_Exitoso.Pass
    GoTo Cleanup

TestFail:
    Test_ObtenerExpedientePorId_Exitoso.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    ' 7. Revertir la transacción para limpiar los datos y cerrar la conexión
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set result = Nothing
    Set repo = Nothing
    Set config = Nothing
    Set db = Nothing
    Set fs = Nothing
End Function

Private Function Test_ObtenerExpedientePorId_NoEncontradoDevuelveNothing() As CTestResult
    Set Test_ObtenerExpedientePorId_NoEncontradoDevuelveNothing = New CTestResult
    Test_ObtenerExpedientePorId_NoEncontradoDevuelveNothing.Initialize "ObtenerExpedientePorId con un ID inexistente debe devolver Nothing"
    
    Dim config As IConfig
    Dim repo As IExpedienteRepository
    Dim result As EExpediente
    Dim db As DAO.Database
    Dim fs As IFileSystem
    
    On Error GoTo TestFail

    ' Arrange: Conectar a la base de datos activa de forma segura
    Set config = modTestContext.GetTestConfig()
    Set repo = modRepositoryFactory.CreateExpedienteRepository(config)
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim dbPath As String: dbPath = config.GetDataPath()
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 100, "Test.Arrange", "La BD de prueba de Expedientes no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath, False, False)
    
    ' Act
    DBEngine.BeginTrans
    Set result = repo.ObtenerExpedientePorId(9999) ' ID заведомо несуществующий

    ' Assert
    modAssert.AssertIsNull result, "El expediente devuelto debe ser Nothing para un ID inexistente."
    
    Test_ObtenerExpedientePorId_NoEncontradoDevuelveNothing.Pass
    GoTo Cleanup

TestFail:
    Test_ObtenerExpedientePorId_NoEncontradoDevuelveNothing.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set result = Nothing
    Set repo = Nothing
    Set config = Nothing
    Set db = Nothing
    Set fs = Nothing
End Function

