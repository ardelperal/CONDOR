Attribute VB_Name = "TIAuthRepository"

Option Compare Database
Option Explicit

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO v2.1 - CORREGIDO)
' ============================================================================

Public Function TIAuthRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIAuthRepository - Pruebas de Integración (CORREGIDO)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestGetUserAuthData_AdminUser_ReturnsCorrectData()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "Error en SuiteSetup/Teardown: (" & Err.Number & ") " & Err.Description & " | Fuente: " & Err.Source
        suiteResult.AddResult errorTest
    End If
    
    Set TIAuthRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    ' Prepara la base de datos de Lanzadera para toda la suite.
    Const LANZADERA_TEMPLATE As String = "Lanzadera_test_template.accdb"
    Const LANZADERA_ACTIVE As String = "Lanzadera_workspace_test.accdb"
    
    Call modTestUtils.PrepareTestDatabase(LANZADERA_TEMPLATE, LANZADERA_ACTIVE)
    
    ' --- INICIO SEEDING ---
    Dim db As DAO.Database, pathDb As String
    pathDb = modTestUtils.GetWorkspacePath() & "Lanzadera_workspace_test.accdb"
    Set db = DBEngine.OpenDatabase(pathDb, False, False, ";PWD=dpddpd")
    Dim email As String: email = "admin@example.com"
    Dim idApp As Long: idApp = 231
    On Error Resume Next
    db.Execute "DELETE FROM TbUsuariosAplicacionesPermisos WHERE CorreoUsuario='" & email & "' AND IDAplicacion=" & idApp, dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE CorreoUsuario='" & email & "'", dbFailOnError
    On Error GoTo 0
    db.Execute "INSERT INTO TbUsuariosAplicaciones (CorreoUsuario, EsAdministrador) VALUES ('" & email & "','Sí')", dbFailOnError
    db.Execute "INSERT INTO TbUsuariosAplicacionesPermisos (CorreoUsuario, IDAplicacion, EsUsuarioAdministrador, EsUsuarioCalidad, EsUsuarioTecnico) " & _
               "VALUES ('" & email & "'," & idApp & ", True, True, True)", dbFailOnError
    db.Close: Set db = Nothing
    ' --- FIN SEEDING ---
End Sub
Private Sub SuiteTeardown()
    ' Limpieza estandarizada a través de la utilidad central.
    Call modTestUtils.CleanupTestDatabase("Lanzadera_workspace_test.accdb")
End Sub

' ============================================================================
' TEST INDIVIDUAL (SIN CAMBIOS)
' ============================================================================

Private Function TestGetUserAuthData_AdminUser_ReturnsCorrectData() As CTestResult
    Set TestGetUserAuthData_AdminUser_ReturnsCorrectData = New CTestResult
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Initialize "GetUserAuthData para Admin debe devolver datos correctos"
    
    Dim repo As IAuthRepository, authData As EAuthData
    Dim fs As IFileSystem
    On Error GoTo TestFail

    ' Arrange: Crear configuración LOCAL apuntando a la BD de prueba de Lanzadera
    Dim config As IConfig
    Dim mockConfigImpl As New CMockConfig
    Dim activeDbPath As String: activeDbPath = modTestUtils.GetWorkspacePath() & "Lanzadera_workspace_test.accdb"
    mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", modTestUtils.GetWorkspacePath() & "Lanzadera_workspace_test.accdb"
    mockConfigImpl.SetSetting "ID_APLICACION_CONDOR", "231"
    mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
    Set config = mockConfigImpl

    ' Crear una instancia REAL del repositorio, inyectando la configuración local
    Set repo = modRepositoryFactory.CreateAuthRepository(config)
    
    ' Arrange: Verificar que la BD existe y conectar
    Set fs = modFileSystemFactory.CreateFileSystem(config)
    Dim dbPath As String: dbPath = modTestUtils.GetWorkspacePath() & "Lanzadera_workspace_test.accdb"

    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 100, "Test.Arrange", "La BD de prueba no existe en la ruta esperada: " & dbPath
    End If
    
    ' Act: Ejecutar el método bajo prueba (los datos ya existen gracias a SuiteSetup)
    Set authData = repo.GetUserAuthData("admin@example.com")
    
    ' Assert: Verificar los resultados
    modAssert.AssertNotNull authData, "El objeto AuthData no debe ser nulo."
    modAssert.AssertTrue authData.UserExists, "UserExists debe ser True."
    modAssert.AssertTrue authData.IsGlobalAdmin, "IsGlobalAdmin debe ser True."
    
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Pass
    GoTo Cleanup

TestFail:
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set fs = Nothing
    Set authData = Nothing: Set repo = Nothing
End Function
