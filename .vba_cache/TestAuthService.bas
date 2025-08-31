Attribute VB_Name = "TestAuthService"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CAuthService
' ============================================================================
' Este módulo contiene pruebas unitarias aisladas para CAuthService
' utilizando mocks para todas las dependencias externas.
' ============================================================================

' Función principal que ejecuta todas las pruebas del módulo
Public Function TestAuthServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TestAuthService")
    
    ' Ejecutar todas las pruebas unitarias
    Call suiteResult.AddTestResult(TestGetUserRoleUsuarioAdministradorDevuelveRolAdministrador())
    Call suiteResult.AddTestResult(TestGetUserRoleUsuarioCalidadDevuelveRolCalidad())
    Call suiteResult.AddTestResult(TestGetUserRoleUsuarioTecnicoDevuelveRolTecnico())
    Call suiteResult.AddTestResult(TestGetUserRoleUsuarioDesconocidoDevuelveRolDesconocido())
    
    Set TestAuthServiceRunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetUserRole
' ============================================================================

' Prueba que GetUserRole devuelve RolAdmin para usuario administrador
Private Function TestGetUserRoleUsuarioAdministradorDevuelveRolAdministrador() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetUserRole debe devolver RolAdmin para un usuario administrador")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar EAuthData para usuario administrador
    Dim authData As New EAuthData
    authData.UserExists = True
    authData.IsGlobalAdmin = True
    
    Call mockAuthRepository.ConfigureMockData(authData)
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    Call authService.Initialize(mockConfig, mockLogger, mockAuthRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim userRole As UserRole
    userRole = authService.GetUserRole("admin@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals RolAdmin, userRole, "El rol devuelto debería ser RolAdmin"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    Set TestGetUserRoleUsuarioAdministradorDevuelveRolAdministrador = testResult
End Function

' Prueba que GetUserRole devuelve RolCalidad para usuario de calidad
Private Function TestGetUserRoleUsuarioCalidadDevuelveRolCalidad() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetUserRole debe devolver RolCalidad para un usuario de calidad")
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar EAuthData para usuario de calidad
    Dim authData As New EAuthData
    authData.UserExists = True
    authData.IsCalidad = True
    
    mockAuthRepository.ConfigureMockData authData
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize mockConfig, mockLogger, mockAuthRepository, mockErrorHandler
    
    ' Act
    Dim userRole As UserRole
    userRole = authService.GetUserRole("calidad@test.com")
    
    ' Assert
    modAssert.AssertEquals RolCalidad, userRole, "El rol devuelto debería ser RolCalidad"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set TestGetUserRoleUsuarioCalidadDevuelveRolCalidad = testResult
End Function

' Prueba que GetUserRole devuelve RolTecnico para usuario técnico
Private Function TestGetUserRoleUsuarioTecnicoDevuelveRolTecnico() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetUserRole debe devolver RolTecnico para un usuario técnico"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar EAuthData para usuario técnico
    Dim authData As New EAuthData
    authData.UserExists = True
    authData.IsTecnico = True
    
    mockAuthRepository.ConfigureMockData authData
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize mockConfig, mockLogger, mockAuthRepository, mockErrorHandler
    
    ' Act
    Dim userRole As UserRole
    userRole = authService.GetUserRole("tecnico@test.com")
    
    ' Assert
    modAssert.AssertEquals RolTecnico, userRole, "El rol devuelto debería ser RolTecnico"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set TestGetUserRoleUsuarioTecnicoDevuelveRolTecnico = testResult
End Function

' Prueba que GetUserRole devuelve RolDesconocido para usuario no encontrado
Private Function TestGetUserRoleUsuarioDesconocidoDevuelveRolDesconocido() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetUserRole debe devolver RolDesconocido para un usuario no encontrado"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar EAuthData para usuario no encontrado
    Dim authData As New EAuthData
    authData.UserExists = False
    
    mockAuthRepository.ConfigureMockData authData
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize mockConfig, mockLogger, mockAuthRepository, mockErrorHandler
    
    ' Act
    Dim userRole As UserRole
    userRole = authService.GetUserRole("desconocido@test.com")
    
    ' Assert
    modAssert.AssertEquals RolDesconocido, userRole, "El rol devuelto debería ser RolDesconocido"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set TestGetUserRoleUsuarioDesconocidoDevuelveRolDesconocido = testResult
End Function














