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
    
    ' Insertar los 7 nuevos estados del flujo refactorizado
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES (1, 'Registrado', 'Estado inicial', TRUE, FALSE);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES (2, 'Desarrollo', 'En desarrollo', FALSE, FALSE);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES (3, 'Modificación', 'Requiere modificaciones', FALSE, FALSE);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES (4, 'Validación', 'En validación', FALSE, FALSE);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES (5, 'Revisión', 'En revisión', FALSE, FALSE);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES (6, 'Formalización', 'En formalización', FALSE, FALSE);", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES (7, 'Aprobada', 'Estado final aprobado', FALSE, TRUE);", dbFailOnError
    
    ' Insertar las transiciones del nuevo flujo
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (1, 1, 2, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (2, 2, 3, 'Ingenieria');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (3, 3, 4, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (4, 3, 5, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (5, 4, 5, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (6, 4, 5, 'Ingenieria');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (7, 5, 6, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (8, 6, 7, 'Calidad');", dbFailOnError
    
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
    
    Dim repo As IWorkflowRepository, db As DAO.Database
    Dim fs As IFileSystem
    Dim dbPath As String
    On Error GoTo TestFail

    ' Arrange: El repositorio obtiene la configuración del contexto centralizado
    Set repo = modRepositoryFactory.CreateWorkflowRepository()
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = modTestUtils.GetProjectPath() & TEST_DB_ACTIVE
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 103, "Test.Arrange", "La BD de prueba de Workflow no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath)
    
    ' Act
    DBEngine.BeginTrans
    Dim isValid As Boolean
    isValid = repo.IsValidTransition("", "Registrado", "Desarrollo")
    
    ' Assert
    modAssert.AssertTrue isValid, "La transición Registrado -> Desarrollo debería ser válida."

    TestIsValidTransition_TrueForValidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_TrueForValidPath.Fail "Error inesperado: " & Err.Description
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set repo = Nothing: Set db = Nothing: Set fs = Nothing
End Function

Private Function TestIsValidTransition_FalseForInvalidPath() As CTestResult
    Set TestIsValidTransition_FalseForInvalidPath = New CTestResult
    TestIsValidTransition_FalseForInvalidPath.Initialize "IsValidTransition debe devolver False para una transición inválida"
    
    Dim repo As IWorkflowRepository, db As DAO.Database
    Dim fs As IFileSystem
    Dim dbPath As String
    On Error GoTo TestFail

    ' Arrange: El repositorio obtiene la configuración del contexto centralizado
    Set repo = modRepositoryFactory.CreateWorkflowRepository()
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = modTestUtils.GetProjectPath() & TEST_DB_ACTIVE
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 103, "Test.Arrange", "La BD de prueba de Workflow no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath)
    
    ' Act
    DBEngine.BeginTrans
    Dim isValid As Boolean
    isValid = repo.IsValidTransition("", "Registrado", "Aprobada")
    
    ' Assert
    modAssert.AssertFalse isValid, "La transición Registrado -> Aprobada (inválida) no debería ser válida."
    
    TestIsValidTransition_FalseForInvalidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_FalseForInvalidPath.Fail "Error inesperado: " & Err.Description
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set repo = Nothing: Set db = Nothing
End Function

Private Function TestGetNextStates_ReturnsCorrectStates() As CTestResult
    Set TestGetNextStates_ReturnsCorrectStates = New CTestResult
    TestGetNextStates_ReturnsCorrectStates.Initialize "GetNextStates debe devolver los estados siguientes correctos"
    
    Dim repo As IWorkflowRepository, db As DAO.Database, nextStates As Scripting.Dictionary
    Dim fs As IFileSystem
    Dim dbPath As String
    On Error GoTo TestFail

    ' Arrange: El repositorio obtiene la configuración del contexto centralizado
    Set repo = modRepositoryFactory.CreateWorkflowRepository()
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = modTestUtils.GetProjectPath() & TEST_DB_ACTIVE
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 103, "Test.Arrange", "La BD de prueba de Workflow no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath)
    
    ' Act
    DBEngine.BeginTrans
    Set nextStates = repo.GetNextStates(1, "Calidad")
    
    ' Assert
    modAssert.AssertEquals 1, nextStates.Count, "Debe haber exactamente un estado siguiente."
    modAssert.AssertTrue nextStates.Exists(2), "El estado siguiente debe ser ID 2 (Desarrollo)."
    
    TestGetNextStates_ReturnsCorrectStates.Pass
    GoTo Cleanup
TestFail:
    TestGetNextStates_ReturnsCorrectStates.Fail "Error inesperado: " & Err.Description
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set nextStates = Nothing: Set repo = Nothing: Set db = Nothing: Set fs = Nothing
End Function