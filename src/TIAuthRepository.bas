Attribute VB_Name = "TIAuthRepository"
Option Compare Database
Option Explicit

' Constantes para el autoaprovisionamiento de la BD de Lanzadera
Private Const LANZADERA_TEMPLATE_PATH As String = "back\test_db\templates\Lanzadera_test_template.accdb"
Private Const LANZADERA_ACTIVE_PATH As String = "back\test_db\active\Lanzadera_integration_test.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (PATRÓN OPTIMIZADO)
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (PATRÓN OPTIMIZADO CON CONFIG LOCAL)
' ============================================================================

Public Function TIAuthRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIAuthRepository - Pruebas de Integración (Config Local)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestGetUserAuthData_AdminUser_ReturnsCorrectData()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIAuthRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & LANZADERA_TEMPLATE_PATH
    Dim activePath As String: activePath = projectPath & LANZADERA_ACTIVE_PATH
    Call modTestUtils.SuiteSetup(templatePath, activePath)
End Sub

Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    Call modTestUtils.SuiteTeardown(activePath)
End Sub

' ============================================================================
' TEST INDIVIDUAL (CON TRANSACCIONES Y AUTO-PROVISIÓN DE DATOS)
' ============================================================================

Private Function TestGetUserAuthData_AdminUser_ReturnsCorrectData() As CTestResult
    Set TestGetUserAuthData_AdminUser_ReturnsCorrectData = New CTestResult
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Initialize "GetUserAuthData para Admin debe devolver datos correctos"
    
    Dim localConfig As IConfig
    Dim repo As IAuthRepository
    Dim authData As EAuthData
    Dim db As DAO.Database
    
    On Error GoTo TestFail

    ' Arrange: 1. Crear una CONFIGURACIÓN LOCAL específica para este test
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
    Set localConfig = mockConfigImpl

    ' Arrange: 2. Inyectar la configuración local en la factoría del repositorio
    Set repo = modRepositoryFactory.CreateAuthRepository(localConfig)
    
    ' Arrange: 3. Conectar a la BD activa y aprovisionar datos
    Set db = DBEngine.OpenDatabase(localConfig.GetLanzaderaDataPath(), False, False, ";PWD=dpddpd")
    DBEngine.BeginTrans
    db.Execute "INSERT INTO TbUsuariosAplicaciones (CorreoUsuario, EsAdministrador) VALUES ('admin@example.com', 'Sí')", dbFailOnError
    
    ' Act: Ejecutar el método bajo prueba
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
    ' Limpieza: Revertir la transacción para eliminar los datos de prueba y cerrar la conexión
    On Error Resume Next
    DBEngine.Rollback
    If Not db Is Nothing Then db.Close
    Set authData = Nothing: Set repo = Nothing: Set localConfig = Nothing: Set db = Nothing
End Function

