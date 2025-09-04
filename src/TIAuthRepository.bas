Attribute VB_Name = "TIAuthRepository"
Option Compare Database
Option Explicit

' Constantes para el autoaprovisionamiento de la BD de Lanzadera
Private Const LANZADERA_TEMPLATE_PATH As String = "back\test_db\templates\Lanzadera_test_template.accdb"
Private Const LANZADERA_ACTIVE_PATH As String = "back\test_db\active\Lanzadera_integration_test.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO v2.0)
' ============================================================================

Public Function TIAuthRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIAuthRepository - Pruebas de Integración (Estándar de Oro)"
    
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
    On Error GoTo ErrorHandler
    ' Aprovisionar la base de datos y los datos maestros para toda la suite
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & LANZADERA_TEMPLATE_PATH
    Dim activePath As String: activePath = projectPath & LANZADERA_ACTIVE_PATH
    Call modTestUtils.SuiteSetup(templatePath, activePath)
    
    ' Insertar los datos de prueba maestros necesarios para esta suite
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(activePath, False, False, ";PWD=dpddpd")
    
    ' ASEGURAR UN ESTADO LIMPIO ANTES DE INSERTAR
    db.Execute "DELETE * FROM TbUsuariosAplicaciones", dbFailOnError
    db.Execute "DELETE * FROM TbUsuariosAplicacionesPermisos", dbFailOnError 'Asegurar limpieza de la segunda tabla
    
    ' INSERTAR EL REGISTRO DE PRUEBA
    db.Execute "INSERT INTO TbUsuariosAplicaciones (Id, CorreoUsuario, EsAdministrador) VALUES (1, 'admin@example.com', 'Sí')", dbFailOnError
    ' INSERTAR EL REGISTRO DE PERMISOS CORRESPONDIENTE
    db.Execute "INSERT INTO TbUsuariosAplicacionesPermisos (CorreoUsuario, IDAplicacion, IDPermiso) VALUES ('admin@example.com', 231, 1)", dbFailOnError
    
    db.Close
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TIAuthRepository.SuiteSetup", Err.Description
End Sub
Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    Call modTestUtils.SuiteTeardown(activePath)
End Sub

' ============================================================================
' TEST INDIVIDUAL (AHORA MÁS LIMPIO)
' ============================================================================

Private Function TestGetUserAuthData_AdminUser_ReturnsCorrectData() As CTestResult
    ' Arrange
    Set TestGetUserAuthData_AdminUser_ReturnsCorrectData = New CTestResult
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Initialize "GetUserAuthData para Admin debe devolver datos correctos"
    
    Dim db As DAO.Database
    Dim repo As IAuthRepository
    Dim authData As EAuthData
    Dim localConfig As IConfig
    Dim activePath As String
    
    On Error GoTo TestFail
    
    activePath = modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    Set db = DBEngine.OpenDatabase(activePath, False, False, ";PWD=dpddpd")
    
    ' --- INICIO DE TRANSACCIÓN ---
    DBEngine.BeginTrans
    
    ' Insertar datos de prueba
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE CorreoUsuario = 'admin@example.com'", dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE Id = 1", dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicacionesPermisos WHERE CorreoUsuario = 'admin@example.com'", dbFailOnError
    db.Execute "INSERT INTO TbUsuariosAplicaciones (Id, CorreoUsuario, EsAdministrador) VALUES (1, 'admin@example.com', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbUsuariosAplicacionesPermisos (CorreoUsuario, IDAplicacion, IDPermiso) VALUES ('admin@example.com', 231, 1)", dbFailOnError
    
    ' Crear configuración local
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", activePath
    mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
    mockConfigImpl.SetSetting "ID_APLICACION_CONDOR", "231"
    Set localConfig = mockConfigImpl
    
    Set repo = modRepositoryFactory.CreateAuthRepository(localConfig)
    
    ' Act
    Set authData = repo.GetUserAuthData("admin@example.com")
    
    ' Assert
    modAssert.AssertNotNull authData, "El objeto AuthData no debe ser nulo."
    modAssert.AssertTrue authData.UserExists, "UserExists debe ser True."
    modAssert.AssertTrue authData.IsGlobalAdmin, "IsGlobalAdmin debe ser True."
    
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Pass

Cleanup:
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    Set repo = Nothing
    Set localConfig = Nothing
    Set authData = Nothing
    
    ' --- REVERSIÓN DE TRANSACCIÓN ---
    ' Esto se ejecuta siempre, haya habido éxito o fallo, para limpiar la BD.
    DBEngine.Rollback
    Exit Function

TestFail:
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Fail "Error: " & Err.Description
    Resume Cleanup
End Function



Private Function GetTestConfig() As IConfig
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
    mockConfigImpl.SetSetting "ID_APLICACION_CONDOR", "231"
    Set GetTestConfig = mockConfigImpl
End Function