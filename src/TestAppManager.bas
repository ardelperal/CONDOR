Attribute VB_Name = "TestAppManager"
Option Compare Database
Option Explicit

Public Function TestAppManagerRunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "TestAppManager"
    
    suite.AddTestResult TestStartApplicationAdminUserSuccess()
    suite.AddTestResult TestStartApplicationUnknownUserFails()
    suite.AddTestResult TestGetCurrentUserRoleReturnsCorrectRole()
    
    Set TestAppManagerRunAll = suite
End Function



Private Function TestStartApplicationAdminUserSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "StartApplication con usuario Admin debe tener éxito y establecer el rol"
    On Error GoTo TestFail
    
    ' Variables locales
    Dim appManager As IAppManager
    Dim mockAuthService As CMockAuthService
    Dim mockConfig As CMockConfig
    Dim mockErrorHandler As CMockErrorHandlerService
    
    ' Arrange
    Set mockAuthService = New CMockAuthService
    mockAuthService.Reset
    Set mockConfig = New CMockConfig
    mockConfig.Reset
    Set mockErrorHandler = New CMockErrorHandlerService
    mockErrorHandler.Reset
    
    ' Simular configuración inicializada
    mockConfig.AddSetting "IsInitialized", True
    
    Dim appManagerImpl As IAppManager
    Set appManagerImpl = New CMockAppManager
    appManagerImpl.Initialize mockAuthService, mockConfig, mockErrorHandler
    Set appManager = appManagerImpl
    
    mockAuthService.ConfigureGetUserRole RolAdmin
    
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
    Set appManager = Nothing
    Set mockAuthService = Nothing
    Set mockConfig = Nothing
    Set mockErrorHandler = Nothing
    Set TestStartApplicationAdminUserSuccess = testResult
End Function

Private Function TestStartApplicationUnknownUserFails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "StartApplication con usuario desconocido debe fallar"
    On Error GoTo TestFail
    
    ' Variables locales
    Dim appManager As IAppManager
    Dim mockAuthService As CMockAuthService
    Dim mockConfig As CMockConfig
    Dim mockErrorHandler As CMockErrorHandlerService
    
    ' Arrange
    Set mockAuthService = New CMockAuthService
    Set mockConfig = New CMockConfig
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' Simular configuración inicializada
    mockConfig.AddSetting "IsInitialized", True
    
    Dim appManagerImpl As IAppManager
    Set appManagerImpl = New CMockAppManager
    appManagerImpl.Initialize mockAuthService, mockConfig, mockErrorHandler
    Set appManager = appManagerImpl
    
    mockAuthService.ConfigureGetUserRole RolDesconocido
    
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
    Set appManager = Nothing
    Set mockAuthService = Nothing
    Set mockConfig = Nothing
    Set mockErrorHandler = Nothing
    Set TestStartApplicationUnknownUserFails = testResult
End Function

Private Function TestGetCurrentUserRoleReturnsCorrectRole() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetCurrentUserRole debe devolver el rol establecido durante el inicio"
    On Error GoTo TestFail
    
    ' Variables locales
    Dim appManager As IAppManager
    Dim mockAuthService As CMockAuthService
    Dim mockConfig As CMockConfig
    Dim mockErrorHandler As CMockErrorHandlerService
    
    ' Arrange
    Set mockAuthService = New CMockAuthService
    Set mockConfig = New CMockConfig
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' Simular configuración inicializada
    mockConfig.AddSetting "IsInitialized", True
    
    Dim appManagerImpl As IAppManager
    Set appManagerImpl = New CMockAppManager
    appManagerImpl.Initialize mockAuthService, mockConfig, mockErrorHandler
    Set appManager = appManagerImpl
    
    mockAuthService.ConfigureGetUserRole RolCalidad
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
    Set appManager = Nothing
    Set mockAuthService = Nothing
    Set mockConfig = Nothing
    Set mockErrorHandler = Nothing
    Set TestGetCurrentUserRoleReturnsCorrectRole = testResult
End Function
















