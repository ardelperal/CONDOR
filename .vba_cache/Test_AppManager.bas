Attribute VB_Name = "Test_AppManager"
Option Compare Database
Option Explicit

' Test_AppManager.bas - Suite de Pruebas Unitarias para modAppManager
' Refactorizado para usar pruebas unitarias aisladas con mocks


#If DEV_MODE Then

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA modAppManager
' ============================================================================

' Función principal de la suite de pruebas
Public Function Test_AppManager_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    Call suite.Initialize("Test_AppManager")
    
    ' Ejecutar todas las pruebas unitarias
    suite.AddTestResult Test_App_Start_AdminUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_App_Start_CalidadUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_App_Start_TecnicoUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_App_Start_DesconocidoUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_Ping_ReturnsPong()
    
    ' Ejecutar prueba de humo
    suite.AddTestResult Test_AppStart_SmokeTest()
    
    Set Test_AppManager_RunAll = suite
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA App_Start CON DIFERENTES ROLES
' ============================================================================

Private Function Test_App_Start_AdminUser_SetsCorrectGlobalRole() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_App_Start_AdminUser_SetsCorrectGlobalRole"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockAuthService As New CMockAuthService
    mockAuthService.SetMockUserRole ROL_ADMINISTRADOR
    modAuthFactory.SetMockAuthService mockAuthService
    
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Act
    modAppManager.App_Start mockErrorHandler, "admin@test.com"
    
    ' Assert
    If modAppManager.g_CurrentUserRole = ROL_ADMINISTRADOR Then
        testResult.Pass
    Else
        testResult.Fail "Expected g_CurrentUserRole to be Rol_Administrador, but was " & modAppManager.g_CurrentUserRole
    End If
    
    ' CleanUp
    modAuthFactory.ResetMock
    
    Set Test_App_Start_AdminUser_SetsCorrectGlobalRole = testResult
    Exit Function
    
TestFail:
    modAuthFactory.ResetMock
    testResult.Fail "Error: " & Err.Number & " - " & Err.Description
    Set Test_App_Start_AdminUser_SetsCorrectGlobalRole = testResult
End Function

Private Function Test_App_Start_CalidadUser_SetsCorrectGlobalRole() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_App_Start_CalidadUser_SetsCorrectGlobalRole"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockAuthService As New CMockAuthService
    mockAuthService.SetMockUserRole Rol_Calidad
    modAuthFactory.SetMockAuthService mockAuthService
    
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Act
    modAppManager.App_Start mockErrorHandler, "calidad@test.com"
    
    ' Assert
    If modAppManager.g_CurrentUserRole = Rol_Calidad Then
        testResult.Pass
    Else
        testResult.Fail "Expected g_CurrentUserRole to be Rol_Calidad, but was " & modAppManager.g_CurrentUserRole
    End If
    
    ' CleanUp
    modAuthFactory.ResetMock
    
    Set Test_App_Start_CalidadUser_SetsCorrectGlobalRole = testResult
    Exit Function
    
TestFail:
    modAuthFactory.ResetMock
    testResult.Fail "Error: " & Err.Number & " - " & Err.Description
    Set Test_App_Start_CalidadUser_SetsCorrectGlobalRole = testResult
End Function

Private Function Test_App_Start_TecnicoUser_SetsCorrectGlobalRole() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_App_Start_TecnicoUser_SetsCorrectGlobalRole"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockAuthService As New CMockAuthService
    mockAuthService.SetMockUserRole Rol_Tecnico
    modAuthFactory.SetMockAuthService mockAuthService
    
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Act
    modAppManager.App_Start mockErrorHandler, "tecnico@test.com"
    
    ' Assert
    If modAppManager.g_CurrentUserRole = Rol_Tecnico Then
        testResult.Pass
    Else
        testResult.Fail "Expected g_CurrentUserRole to be Rol_Tecnico, but was " & modAppManager.g_CurrentUserRole
    End If
    
    ' CleanUp
    modAuthFactory.ResetMock
    
    Set Test_App_Start_TecnicoUser_SetsCorrectGlobalRole = testResult
    Exit Function
    
TestFail:
    modAuthFactory.ResetMock
    testResult.Fail "Error: " & Err.Number & " - " & Err.Description
    Set Test_App_Start_TecnicoUser_SetsCorrectGlobalRole = testResult
End Function

Private Function Test_App_Start_DesconocidoUser_SetsCorrectGlobalRole() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_App_Start_DesconocidoUser_SetsCorrectGlobalRole"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockAuthService As New CMockAuthService
    mockAuthService.SetMockUserRole Rol_Desconocido
    modAuthFactory.SetMockAuthService mockAuthService
    
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Act
    modAppManager.App_Start mockErrorHandler, "desconocido@test.com"
    
    ' Assert
    If modAppManager.g_CurrentUserRole = Rol_Desconocido Then
        testResult.Pass
    Else
        testResult.Fail "Expected g_CurrentUserRole to be Rol_Desconocido, but was " & modAppManager.g_CurrentUserRole
    End If
    
    ' CleanUp
    modAuthFactory.ResetMock
    
    Set Test_App_Start_DesconocidoUser_SetsCorrectGlobalRole = testResult
    Exit Function
    
TestFail:
    modAuthFactory.ResetMock
    testResult.Fail "Error: " & Err.Number & " - " & Err.Description
    Set Test_App_Start_DesconocidoUser_SetsCorrectGlobalRole = testResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA FUNCIÓN Ping
' ============================================================================

Private Function Test_Ping_ReturnsPong() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_Ping_ReturnsPong"
    
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim result As String
    result = modAppManager.Ping()
    
    ' Assert
    If result = "Pong" Then
        testResult.Pass
    Else
        testResult.Fail "Expected 'Pong', but got '" & result & "'"
    End If
    
    Set Test_Ping_ReturnsPong = testResult
    Exit Function
    
TestFail:
    testResult.Fail "Error: " & Err.Number & " - " & Err.Description
    Set Test_Ping_ReturnsPong = testResult
End Function

' ============================================================================
' PRUEBA DE HUMO PARA App_Start CON INSTANCIAS REALES
' ============================================================================

Private Function Test_AppStart_SmokeTest() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AppStart_SmokeTest"
    
    On Error GoTo TestFail
    
    ' Arrange - Crear instancias reales del CErrorHandlerService y sus dependencias
    Dim config As New CConfig
    Dim fileSystem As New CFileSystem
    Dim errorHandler As New CErrorHandlerService
    errorHandler.Initialize config, fileSystem
    
    ' Act - Llamar al procedimiento App_Start con errorHandler real y email de prueba
    modAppManager.App_Start errorHandler, "admin@test.com"
    
    ' Assert - Verificar que no se lancen errores y que g_CurrentUserRole se establezca correctamente
    If modAppManager.g_CurrentUserRole = ROL_ADMINISTRADOR Then
        testResult.Pass
    Else
        testResult.Fail "Expected g_CurrentUserRole to be ROL_ADMINISTRADOR, but was " & modAppManager.g_CurrentUserRole
    End If
    
    Set Test_AppStart_SmokeTest = testResult
    Exit Function
    
TestFail:
    testResult.Fail "Error during smoke test: " & Err.Number & " - " & Err.Description
    Set Test_AppStart_SmokeTest = testResult
End Function

#End If
















