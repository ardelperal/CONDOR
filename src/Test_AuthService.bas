Attribute VB_Name = "Test_AuthService"
Option Compare Database
Option Explicit


#If DEV_MODE Then

' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CAuthService
' ============================================================================
' Este módulo contiene pruebas unitarias aisladas para CAuthService
' utilizando mocks para todas las dependencias externas.
' Sigue la Lección 10: El Aislamiento de las Pruebas Unitarias con Mocks no es Negociable

' Función principal que ejecuta todas las pruebas del módulo
Public Function Test_AuthService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("Test_AuthService")
    
    ' Ejecutar todas las pruebas unitarias
    Call suiteResult.AddTestResult(Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador())
    Call suiteResult.AddTestResult(Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad())
    Call suiteResult.AddTestResult(Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico())
    Call suiteResult.AddTestResult(Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido())
    
    Set Test_AuthService_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetUserRole
' ============================================================================

' Prueba que GetUserRole devuelve Rol_Administrador para usuario administrador
Private Function Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador")
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim mockConfig As New CMockConfig
    Call mockConfig.AddSetting("BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb")
    Call mockConfig.AddSetting("DATABASE_PASSWORD", "testpassword")
    
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar T_AuthData para usuario administrador
    Dim authData As New T_AuthData
    authData.UserExists = True
    authData.IsGlobalAdmin = True
    authData.IsAppAdmin = False
    authData.IsCalidad = False
    authData.IsTecnico = False
    
    Call mockAuthRepository.ConfigureMockData(authData)
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    Call authService.Initialize(mockConfig, mockLogger, mockAuthRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("admin@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals ROL_ADMINISTRADOR, userRole, "GetUserRole debe devolver Rol_Administrador para usuario administrador"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    Set authData = Nothing
    mockConfig.Reset
    mockLogger.Reset
    mockAuthRepository.Reset
    mockErrorHandler.Reset
    Set Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador = testResult
    Exit Function
End Function

' Prueba que GetUserRole devuelve Rol_Calidad para usuario de calidad
Private Function Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad")
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim mockConfig As New CMockConfig
    Call mockConfig.AddSetting("BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb")
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar T_AuthData para usuario de calidad
    Dim authData As New T_AuthData
    authData.UserExists = True
    authData.IsGlobalAdmin = False
    authData.IsAppAdmin = False
    authData.IsCalidad = True
    authData.IsTecnico = False
    
    mockAuthRepository.ConfigureMockData authData
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize mockConfig, mockLogger, mockAuthRepository, mockErrorHandler
    
    ' Act - Ejecutar el método bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("calidad@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals Rol_Calidad, userRole, "GetUserRole debe devolver Rol_Calidad para usuario de calidad"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set authData = Nothing
    mockConfig.Reset
    mockLogger.Reset
    mockAuthRepository.Reset
    mockErrorHandler.Reset
    Set Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad = testResult
    Exit Function
End Function

' Prueba que GetUserRole devuelve Rol_Tecnico para usuario técnico
Private Function Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar T_AuthData para usuario técnico
    Dim authData As New T_AuthData
    authData.UserExists = True
    authData.IsGlobalAdmin = False
    authData.IsAppAdmin = False
    authData.IsCalidad = False
    authData.IsTecnico = True
    
    mockAuthRepository.ConfigureMockData authData
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize mockConfig, mockLogger, mockAuthRepository, mockErrorHandler
    
    ' Act - Ejecutar el método bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("tecnico@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals Rol_Tecnico, userRole, "GetUserRole debe devolver Rol_Tecnico para usuario técnico"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set authData = Nothing
    mockConfig.Reset
    mockLogger.Reset
    mockAuthRepository.Reset
    mockErrorHandler.Reset
    Set Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico = testResult
    Exit Function
End Function

' Prueba que GetUserRole devuelve Rol_Desconocido para usuario no encontrado
Private Function Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    Dim mockAuthRepository As New CMockAuthRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Configurar T_AuthData para usuario no encontrado
    Dim authData As New T_AuthData
    authData.UserExists = False
    authData.IsGlobalAdmin = False
    authData.IsAppAdmin = False
    authData.IsCalidad = False
    authData.IsTecnico = False
    
    mockAuthRepository.ConfigureMockData authData
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize mockConfig, mockLogger, mockAuthRepository, mockErrorHandler
    
    ' Act - Ejecutar el método bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("inexistente@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals Rol_Desconocido, userRole, "GetUserRole debe devolver Rol_Desconocido para usuario no encontrado"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set authData = Nothing
    mockConfig.Reset
    mockLogger.Reset
    mockAuthRepository.Reset
    mockErrorHandler.Reset
    Set Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido = testResult
    Exit Function
End Function

#End If














