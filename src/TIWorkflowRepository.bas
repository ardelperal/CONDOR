Attribute VB_Name = "TIWorkflowRepository"
Option Compare Database
Option Explicit

Private Const TEST_DB_TEMPLATE As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const TEST_DB_ACTIVE As String = "back\test_db\active\workflow_test.accdb"

Private m_Config As IConfig
Private m_Repo As IWorkflowRepository

Public Function TIWorkflowRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWorkflowRepository"
    
    suiteResult.AddResult TestIsValidTransition_TrueForValidPath()
    suiteResult.AddResult TestIsValidTransition_FalseForInvalidPath()
    suiteResult.AddResult TestGetNextStates_ReturnsCorrectStates()
    
    Set TIWorkflowRepositoryRunAll = suiteResult
End Function

Private Sub Setup()
    On Error GoTo TestFail
    
    Dim testDbPath As String: testDbPath = modTestUtils.GetProjectPath & TEST_DB_ACTIVE
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath & TEST_DB_TEMPLATE, testDbPath
    
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", testDbPath
    mockConfigImpl.SetSetting "DATABASE_PASSWORD", "dpddpd"
    Set m_Config = mockConfigImpl
    
    Call PopulateTestData(m_Config)
    
    Set m_Repo = modRepositoryFactory.CreateWorkflowRepository(m_Config)
    Exit Sub
TestFail:
    Err.Raise Err.Number, "TIWorkflowRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Set m_Repo = Nothing
    Set m_Config = Nothing
    
    ' Lógica de limpieza correcta usando IFileSystem
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testDbPath As String
    testDbPath = modTestUtils.GetProjectPath & TEST_DB_ACTIVE
    
    If fs.FileExists(testDbPath) Then
        fs.DeleteFile testDbPath, True ' Forzar borrado
    End If
    
    Set fs = Nothing
End Sub

Private Sub PopulateTestData(ByVal config As IConfig)
    Dim db As DAO.Database
    Dim dbPath As String: dbPath = config.GetDataPath()
    Dim dbPass As String: dbPass = config.GetDatabasePassword()
    Set db = DBEngine.OpenDatabase(dbPath, False, False, ";PWD=" & dbPass)
    
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, esEstadoInicial, esEstadoFinal) VALUES (1, 'Borrador', True, False);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, esEstadoInicial, esEstadoFinal) VALUES (2, 'Aprobado', False, True);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, esEstadoInicial, esEstadoFinal) VALUES (3, 'Rechazado', False, True);", dbFailOnError
    
    db.Execute "INSERT INTO tbTransiciones (idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (1, 2, 'Admin');", dbFailOnError
    
    db.Close
    Set db = Nothing
End Sub

Private Function TestIsValidTransition_TrueForValidPath() As CTestResult
    Set TestIsValidTransition_TrueForValidPath = New CTestResult
    TestIsValidTransition_TrueForValidPath.Initialize "IsValidTransition debe devolver True para una transición válida"
    On Error GoTo TestFail
    Call Setup
    
    Dim isValid As Boolean
    isValid = m_Repo.IsValidTransition(1, 2, "Admin")
    
    modAssert.AssertTrue isValid, "La transición 1 -> 2 para Admin debería ser válida."
    
    TestIsValidTransition_TrueForValidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_TrueForValidPath.Fail Err.Description
Cleanup:
    Call Teardown
End Function

Private Function TestIsValidTransition_FalseForInvalidPath() As CTestResult
    Set TestIsValidTransition_FalseForInvalidPath = New CTestResult
    TestIsValidTransition_FalseForInvalidPath.Initialize "IsValidTransition debe devolver False para una transición inválida"
    On Error GoTo TestFail
    Call Setup
    
    Dim isValid As Boolean
    isValid = m_Repo.IsValidTransition(1, 3, "Admin")
    
    modAssert.AssertFalse isValid, "La transición 1 -> 3 no debería ser válida."
    
    TestIsValidTransition_FalseForInvalidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_FalseForInvalidPath.Fail Err.Description
Cleanup:
    Call Teardown
End Function

Private Function TestGetNextStates_ReturnsCorrectStates() As CTestResult
    Set TestGetNextStates_ReturnsCorrectStates = New CTestResult
    TestGetNextStates_ReturnsCorrectStates.Initialize "GetNextStates debe devolver los estados siguientes correctos"
    Dim nextStates As Scripting.Dictionary
    On Error GoTo TestFail
    Call Setup
    
    Set nextStates = m_Repo.GetNextStates(1, "Admin")
    
    modAssert.AssertEquals 1, nextStates.Count, "Debe haber exactamente un estado siguiente."
    modAssert.AssertTrue nextStates.Exists(2), "El estado siguiente debe ser ID 2 (Aprobado)."
    
    TestGetNextStates_ReturnsCorrectStates.Pass
    GoTo Cleanup
TestFail:
    TestGetNextStates_ReturnsCorrectStates.Fail Err.Description
Cleanup:
    Set nextStates = Nothing
    Call Teardown
End Function


