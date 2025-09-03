Attribute VB_Name = "TIAuthRepository"
Option Compare Database
Option Explicit

' Constantes para el autoaprovisionamiento de la BD de Lanzadera
Private Const LANZADERA_TEMPLATE_PATH As String = "back\test_db\templates\Lanzadera_test_template.accdb"
Private Const LANZADERA_ACTIVE_PATH As String = "back\test_db\active\Lanzadera_integration_test.accdb"

Private m_Config As IConfig
Private m_Repo As IAuthRepository

Public Function TIAuthRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIAuthRepository"
    
    suiteResult.AddResult TestGetUserAuthData_AdminUser_ReturnsCorrectData()
    
    Set TIAuthRepositoryRunAll = suiteResult
End Function

Private Sub Setup()
    On Error GoTo TestFail
    
    ' 1. Preparar la BD de prueba de LANZADERA
    Dim testDbPath As String: testDbPath = modTestUtils.GetProjectPath & LANZADERA_ACTIVE_PATH
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath & LANZADERA_TEMPLATE_PATH, testDbPath
    
    ' 2. Crear y configurar el MOCK de configuración
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "LOG_FILE_PATH", modTestUtils.GetProjectPath & "condor_test_run.log"
    mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", testDbPath
    mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
    Set m_Config = mockConfigImpl
    
    ' 3. Instanciar el Repositorio real inyectando la configuración de prueba
    Set m_Repo = modRepositoryFactory.CreateAuthRepository(m_Config)
    Exit Sub
    
TestFail:
    Err.Raise Err.Number, "TIAuthRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Set m_Repo = Nothing
    Set m_Config = Nothing
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFile modTestUtils.GetProjectPath & LANZADERA_ACTIVE_PATH, True
    Set fs = Nothing
End Sub

Private Function TestGetUserAuthData_AdminUser_ReturnsCorrectData() As CTestResult
    Set TestGetUserAuthData_AdminUser_ReturnsCorrectData = New CTestResult
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Initialize "GetUserAuthData para Admin debe devolver datos correctos"
    
    Dim authData As EAuthData
    On Error GoTo TestFail
    Call Setup
    
    ' Act
    ' La base de datos de prueba contiene un usuario admin@example.com
    Set authData = m_Repo.GetUserAuthData("admin@example.com")
    
    ' Assert
    modAssert.AssertNotNull authData, "El objeto AuthData no debe ser nulo."
    modAssert.AssertTrue authData.UserExists, "El usuario admin debería existir en la BD de prueba."
    modAssert.AssertTrue authData.IsGlobalAdmin, "El usuario admin debería ser administrador global."
    
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Pass
    GoTo Cleanup
    
TestFail:
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
    Set authData = Nothing
End Function

