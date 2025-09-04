Attribute VB_Name = "TIAuthRepository"

Option Compare Database
Option Explicit

' Constantes para el autoaprovisionamiento de la BD de Lanzadera
Private Const LANZADERA_TEMPLATE_PATH As String = "back\test_db\templates\Lanzadera_test_template.accdb"
Private Const LANZADERA_ACTIVE_PATH As String = "back\test_db\active\Lanzadera_integration_test.accdb"

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
    On Error GoTo errorHandler
    ' Aprovisionar la base de datos y los datos maestros para toda la suite
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & LANZADERA_TEMPLATE_PATH
    Dim activePath As String: activePath = projectPath & LANZADERA_ACTIVE_PATH
    Call modTestUtils.SuiteSetup(templatePath, activePath)
    
    ' Insertar los datos de prueba maestros necesarios para esta suite
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(activePath, False, False, ";PWD=dpddpd")
    
    ' ASEGURAR UN ESTADO LIMPIO RESPETANDO LA INTEGRIDAD REFERENCIAL
    db.Execute "DELETE * FROM TbUsuariosAplicacionesPermisos", dbFailOnError
    db.Execute "DELETE * FROM TbUsuariosAplicaciones", dbFailOnError
    
    ' INSERTAR EL REGISTRO DE PRUEBA EN AMBAS TABLAS
    db.Execute "INSERT INTO TbUsuariosAplicaciones (Id, CorreoUsuario, EsAdministrador, Password) VALUES (1, 'admin@example.com', 'Sí', 'password')", dbFailOnError
    ' INSERTAR EL REGISTRO DE PERMISOS CON EL ESQUEMA CORRECTO
    db.Execute "INSERT INTO TbUsuariosAplicacionesPermisos (CorreoUsuario, IDAplicacion, EsUsuarioAdministrador) VALUES ('admin@example.com', 231, True)", dbFailOnError

    db.Close
    Set db = Nothing
    Exit Sub
errorHandler:
    If Not db Is Nothing Then db.Close
    Err.Raise Err.Number, "TIAuthRepository.SuiteSetup", Err.Description
End Sub
Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    Call modTestUtils.SuiteTeardown(activePath)
End Sub

' ============================================================================
' TEST INDIVIDUAL (SIN CAMBIOS)
' ============================================================================

Private Function TestGetUserAuthData_AdminUser_ReturnsCorrectData() As CTestResult
    Set TestGetUserAuthData_AdminUser_ReturnsCorrectData = New CTestResult
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Initialize "GetUserAuthData para Admin debe devolver datos correctos"
    
    Dim localConfig As IConfig, repo As IAuthRepository, authData As EAuthData
    Dim fs As IFileSystem
    On Error GoTo TestFail

    ' Arrange: 1. Crear una CONFIGURACIÓN LOCAL específica para este test
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
    mockConfigImpl.SetSetting "ID_APLICACION_CONDOR", "231"
    Set localConfig = mockConfigImpl

    ' Arrange: 2. Inyectar la configuración local en la factoría del repositorio
    Set repo = modRepositoryFactory.CreateAuthRepository(localConfig)
    
    ' Arrange: 3. Verificar que la BD existe y conectar
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim dbPath As String: dbPath = localConfig.GetLanzaderaDataPath()

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
    Set authData = Nothing: Set repo = Nothing: Set localConfig = Nothing
End Function
