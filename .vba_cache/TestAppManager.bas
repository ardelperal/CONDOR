Attribute VB_Name = "TestAppManager"
Option Compare Database
Option Explicit

Private appManager As IAppManager
Private mockAuthService As CMockAuthService
Private mockConfig As CMockConfig
Private mockErrorHandler As CMockErrorHandlerService

Public Function TestAppManagerRunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "TestAppManager"
    
    suite.AddTestResult TestStartApplicationAdminUserSuccess()
    suite.AddTestResult TestStartApplicationUnknownUserFails()
    suite.AddTestResult TestGetCurrentUserRoleReturnsCorrectRole()
    
    Set TestAppManagerRunAll = suite
End Function

Private Sub Setup()
    Set mockAuthService = New CMockAuthService
    Set mockConfig = New CMockConfig
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' Simular configuración inicializada
    mockConfig.AddSetting "IsInitialized", True
    
    Dim appManagerImpl As New CAppManager
    appManagerImpl.Initialize mockAuthService, mockConfig, mockErrorHandler
    Set appManager = appManagerImpl
End Sub

Private Sub Teardown()
    Set appManager = Nothing
    Set mockAuthService = Nothing
    Set mockConfig = Nothing
    Set mockErrorHandler = Nothing
End Sub

Private Function TestStartApplicationAdminUserSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "StartApplication con usuario Admin debe tener éxito y establecer el rol"
    On Error GoTo TestFail
    
    ' Arrange
    Setup
    mockAuthService.SetMockUserRole RolAdmin
    
    ' Act
    Dim success As Boolean
    success = appManager.StartApplication("admin@test.com")
    
    ' Assert
    modAssert.AssertTrue success, "La aplicación debería iniciarse correctamente."
    modAssert.AssertEquals RolAdmin, appManager.GetCurrentUserRole(), "El rol del usuario debe ser Admin."
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestStartApplicationAdminUserSuccess = testResult
End Function

Private Function TestStartApplicationUnknownUserFails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "StartApplication con usuario desconocido debe fallar"
    On Error GoTo TestFail
    
    ' Arrange
    Setup
    mockAuthService.SetMockUserRole RolDesconocido
    
    ' Act
    Dim success As Boolean
    success = appManager.StartApplication("unknown@test.com")
    
    ' Assert
    modAssert.AssertFalse success, "El inicio de la aplicación debería fallar."
    modAssert.AssertEquals RolDesconocido, appManager.GetCurrentUserRole(), "El rol debe permanecer como Desconocido."
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestStartApplicationUnknownUserFails = testResult
End Function

Private Function TestGetCurrentUserRoleReturnsCorrectRole() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetCurrentUserRole debe devolver el rol establecido durante el inicio"
    On Error GoTo TestFail
    
    ' Arrange
    Setup
    mockAuthService.SetMockUserRole RolCalidad
    appManager.StartApplication "calidad@test.com"
    
    ' Act
    Dim role As UserRole
    role = appManager.GetCurrentUserRole()
    
    ' Assert
    modAssert.AssertEquals RolCalidad, role, "Debe devolver el rol de Calidad."
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestGetCurrentUserRoleReturnsCorrectRole = testResult
End Function
















