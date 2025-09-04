Attribute VB_Name = "TIWorkflowRepository"
Option Compare Database
Option Explicit

Private Const TEST_DB_TEMPLATE As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const TEST_DB_ACTIVE As String = "back\test_db\active\workflow_test.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TIWorkflowRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWorkflowRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestIsValidTransition_TrueForValidPath()
    suiteResult.AddResult TestIsValidTransition_FalseForInvalidPath()
    suiteResult.AddResult TestGetNextStates_ReturnsCorrectStates()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIWorkflowRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler
    ' Aprovisionar la base de datos y los datos maestros para toda la suite
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & TEST_DB_TEMPLATE
    Dim activePath As String: activePath = projectPath & TEST_DB_ACTIVE
    Call modTestUtils.SuiteSetup(templatePath, activePath)
    
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(activePath)
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (1, 'Borrador');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (2, 'Aprobado');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (3, 'Rechazado');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (1, 2, 'Admin');", dbFailOnError
    db.Close
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TIWorkflowRepository.SuiteSetup", Err.Description
End Sub

Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & TEST_DB_ACTIVE
    Call modTestUtils.SuiteTeardown(activePath)
End Sub

' ============================================================================
' TESTS INDIVIDUALES (CON TRANSACCIONES)
' ============================================================================

Private Function TestIsValidTransition_TrueForValidPath() As CTestResult
    Set TestIsValidTransition_TrueForValidPath = New CTestResult
    TestIsValidTransition_TrueForValidPath.Initialize "IsValidTransition debe devolver True para una transición válida"
    
    Dim localConfig As IConfig, repo As IWorkflowRepository, db As DAO.Database
    On Error GoTo TestFail

    ' Arrange
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", modTestUtils.GetProjectPath() & TEST_DB_ACTIVE
    Set localConfig = mockConfigImpl
    Set repo = modRepositoryFactory.CreateWorkflowRepository(localConfig)
    Set db = DBEngine.OpenDatabase(localConfig.GetDataPath())
    
    ' Act
    DBEngine.BeginTrans
    Dim isValid As Boolean
    isValid = repo.IsValidTransition(1, 2, "Admin")
    
    ' Assert
    modAssert.AssertTrue isValid, "La transición 1 -> 2 para Admin debería ser válida."

    TestIsValidTransition_TrueForValidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_TrueForValidPath.Fail "Error inesperado: " & Err.Description
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set repo = Nothing: Set localConfig = Nothing: Set db = Nothing
End Function

Private Function TestIsValidTransition_FalseForInvalidPath() As CTestResult
    Set TestIsValidTransition_FalseForInvalidPath = New CTestResult
    TestIsValidTransition_FalseForInvalidPath.Initialize "IsValidTransition debe devolver False para una transición inválida"
    
    Dim localConfig As IConfig, repo As IWorkflowRepository, db As DAO.Database
    On Error GoTo TestFail

    ' Arrange
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", modTestUtils.GetProjectPath() & TEST_DB_ACTIVE
    Set localConfig = mockConfigImpl
    Set repo = modRepositoryFactory.CreateWorkflowRepository(localConfig)
    Set db = DBEngine.OpenDatabase(localConfig.GetDataPath())
    
    ' Act
    DBEngine.BeginTrans
    Dim isValid As Boolean
    isValid = repo.IsValidTransition(1, 3, "Admin")
    
    ' Assert
    modAssert.AssertFalse isValid, "La transición 1 -> 3 (inválida) no debería ser válida."
    
    TestIsValidTransition_FalseForInvalidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_FalseForInvalidPath.Fail "Error inesperado: " & Err.Description
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set repo = Nothing: Set localConfig = Nothing: Set db = Nothing
End Function

Private Function TestGetNextStates_ReturnsCorrectStates() As CTestResult
    Set TestGetNextStates_ReturnsCorrectStates = New CTestResult
    TestGetNextStates_ReturnsCorrectStates.Initialize "GetNextStates debe devolver los estados siguientes correctos"
    
    Dim localConfig As IConfig, repo As IWorkflowRepository, db As DAO.Database, nextStates As Scripting.Dictionary
    On Error GoTo TestFail

    ' Arrange
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", modTestUtils.GetProjectPath() & TEST_DB_ACTIVE
    Set localConfig = mockConfigImpl
    Set repo = modRepositoryFactory.CreateWorkflowRepository(localConfig)
    Set db = DBEngine.OpenDatabase(localConfig.GetDataPath())
    
    ' Act
    DBEngine.BeginTrans
    Set nextStates = repo.GetNextStates(1, "Admin")
    
    ' Assert
    modAssert.AssertEquals 1, nextStates.Count, "Debe haber exactamente un estado siguiente."
    modAssert.AssertTrue nextStates.Exists(2), "El estado siguiente debe ser ID 2 (Aprobado)."
    
    TestGetNextStates_ReturnsCorrectStates.Pass
    GoTo Cleanup
TestFail:
    TestGetNextStates_ReturnsCorrectStates.Fail "Error inesperado: " & Err.Description
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set nextStates = Nothing: Set repo = Nothing: Set localConfig = Nothing: Set db = Nothing
End Function