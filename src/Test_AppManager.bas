' Test_AppManager.bas - Suite de Pruebas Unitarias para modAppManager
' Refactorizado para usar pruebas unitarias aisladas con mocks

Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA modAppManager
' ============================================================================

' Función principal de la suite de pruebas
Public Function Test_AppManager_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "Test_AppManager"
    
    ' Ejecutar todas las pruebas unitarias
    suite.AddTestResult Test_App_Start_AdminUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_App_Start_CalidadUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_App_Start_TecnicoUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_App_Start_DesconocidoUser_SetsCorrectGlobalRole()
    suite.AddTestResult Test_Ping_ReturnsPong()
    
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
    mockAuthService.SetMockUserRole Rol_Admin
    modAuthFactory.SetMockAuthService mockAuthService
    
    ' Act
    modAppManager.App_Start
    
    ' Assert
    If modAppManager.g_CurrentUserRole = Rol_Admin Then
        testResult.Pass
    Else
        testResult.Fail "Expected g_CurrentUserRole to be Rol_Admin, but was " & modAppManager.g_CurrentUserRole
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
    
    ' Act
    modAppManager.App_Start
    
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
    
    ' Act
    modAppManager.App_Start
    
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
    
    ' Act
    modAppManager.App_Start
    
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

#End If














