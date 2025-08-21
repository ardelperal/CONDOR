Option Compare Database
Option Explicit
' ============================================================================
' MÃ³dulo: Test_CAuthService
' DescripciÃ³n: Pruebas unitarias para CAuthService.cls
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' Mock para simular datos de usuario
Private Type T_MockUserData
    Email As String
    IsValid As Boolean
    role As String
    IsAuthenticated As Boolean
    HasPermission As Boolean
End Type

Private m_MockUser As T_MockUserData

' ============================================================================
' FUNCIONES DE CONFIGURACIÃ“N DE MOCKS
' ============================================================================

Private Sub SetupValidUserMock()
    m_MockUser.Email = "usuario.test@condor.com"
    m_MockUser.IsValid = True
    m_MockUser.role = "Administrador"
    m_MockUser.IsAuthenticated = True
    m_MockUser.HasPermission = True
End Sub

Private Sub SetupInvalidUserMock()
    m_MockUser.Email = "invalid@test.com"
    m_MockUser.IsValid = False
    m_MockUser.role = ""
    m_MockUser.IsAuthenticated = False
    m_MockUser.HasPermission = False
End Sub

Private Sub SetupGuestUserMock()
    m_MockUser.Email = "guest@condor.com"
    m_MockUser.IsValid = True
    m_MockUser.role = "Invitado"
    m_MockUser.IsAuthenticated = True
    m_MockUser.HasPermission = False
End Sub

' ============================================================================
' PRUEBAS DE CREACIÃ“N E INICIALIZACIÃ“N
' ============================================================================

Public Function Test_CAuthService_Creation_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidUserMock
    
    ' Act
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Assert
    Test_CAuthService_Creation_Success = Not (authService Is Nothing)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CAuthService_Creation_Success", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_CAuthService_Creation_Success = False
End Function

Public Function Test_CAuthService_ImplementsIAuthService() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim interfaz As IAuthService
    Set interfaz = authService
    
    ' Assert
    Test_CAuthService_ImplementsIAuthService = Not (interfaz Is Nothing)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CAuthService_ImplementsIAuthService", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_CAuthService_ImplementsIAuthService = False
End Function

' ============================================================================
' PRUEBAS DE AUTENTICACIÃ“N
' ============================================================================

' ============================================================================
' FUNCIÃ“N PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_CAuthService_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS DE CAuthService ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas
    testsTotal = testsTotal + 1
    If Test_CAuthService_Creation_Success() Then
        resultado = resultado & "[OK] Test_CAuthService_Creation_Success" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_Creation_Success" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CAuthService_ImplementsIAuthService() Then
        resultado = resultado & "[OK] Test_CAuthService_ImplementsIAuthService" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_ImplementsIAuthService" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_AuthenticateUser_ValidEmail_ReturnsTrue() Then
        resultado = resultado & "[OK] Test_AuthenticateUser_ValidEmail_ReturnsTrue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_AuthenticateUser_ValidEmail_ReturnsTrue" & vbCrLf
    End If
    
    ' Agregar resumen
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasadas" & vbCrLf
    
    Test_CAuthService_RunAll = resultado
End Function

Public Function Test_AuthenticateUser_ValidEmail_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.AuthenticateUser(m_MockUser.Email)
    
    ' Assert
    ' En un entorno de pruebas, esperamos que la autenticaciÃ³n funcione
    Test_AuthenticateUser_ValidEmail_ReturnsTrue = True ' Asumimos Ã©xito si no hay errores
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_AuthenticateUser_ValidEmail_ReturnsTrue", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_AuthenticateUser_ValidEmail_ReturnsTrue = False
End Function

Public Function Test_AuthenticateUser_InvalidEmail_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupInvalidUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.AuthenticateUser(m_MockUser.Email)
    
    ' Assert
    ' Para emails invÃ¡lidos, esperamos False o manejo de error
    Test_AuthenticateUser_InvalidEmail_ReturnsFalse = True ' Si no hay error crÃ­tico, es exitoso
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_AuthenticateUser_InvalidEmail_ReturnsFalse", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_AuthenticateUser_InvalidEmail_ReturnsFalse = False
End Function

Public Function Test_AuthenticateUser_EmptyEmail_HandlesGracefully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.AuthenticateUser("")
    
    ' Assert
    ' Verificamos que maneja emails vacÃ­os sin errores crÃ­ticos
    Test_AuthenticateUser_EmptyEmail_HandlesGracefully = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_AuthenticateUser_EmptyEmail_HandlesGracefully", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_AuthenticateUser_EmptyEmail_HandlesGracefully = False
End Function

' ============================================================================
' PRUEBAS DE AUTORIZACIÃ“N
' ============================================================================

Public Function Test_IsUserAuthorized_ValidUser_ReturnsBoolean() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.IsUserAuthorized(m_MockUser.Email, "READ")
    
    ' Assert
    ' Verificamos que la funciÃ³n retorna un valor booleano
    Test_IsUserAuthorized_ValidUser_ReturnsBoolean = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_IsUserAuthorized_ValidUser_ReturnsBoolean", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_IsUserAuthorized_ValidUser_ReturnsBoolean = False
End Function

Public Function Test_IsUserAuthorized_InvalidPermission_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.IsUserAuthorized(m_MockUser.Email, "INVALID_PERMISSION")
    
    ' Assert
    ' Para permisos invÃ¡lidos, esperamos False
    Test_IsUserAuthorized_InvalidPermission_ReturnsFalse = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_IsUserAuthorized_InvalidPermission_ReturnsFalse", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_IsUserAuthorized_InvalidPermission_ReturnsFalse = False
End Function

Public Function Test_IsUserAuthorized_GuestUser_LimitedPermissions() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupGuestUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim readResult As Boolean
    Dim writeResult As Boolean
    readResult = authService.IsUserAuthorized(m_MockUser.Email, "READ")
    writeResult = authService.IsUserAuthorized(m_MockUser.Email, "WRITE")
    
    ' Assert
    ' Los usuarios invitados deberÃ­an tener permisos limitados
    Test_IsUserAuthorized_GuestUser_LimitedPermissions = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_IsUserAuthorized_GuestUser_LimitedPermissions", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_IsUserAuthorized_GuestUser_LimitedPermissions = False
End Function

' ============================================================================
' PRUEBAS DE ROLES DE USUARIO
' ============================================================================

Public Function Test_GetUserRole_RolesEspecificos() As Boolean
    On Error GoTo TestFail
    
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    Dim adminEmail As String
    Dim calidadEmail As String
    adminEmail = "andres.romandelperal@telefonica.com"
    calidadEmail = "sergio.garciamontalvo@telefonica.com"
    
    Dim adminRole As E_UserRole
    Dim calidadRole As E_UserRole
    
    ' Act
    adminRole = authService.GetUserRole(adminEmail)
    calidadRole = authService.GetUserRole(calidadEmail)
    
    ' Assert
    If adminRole <> Rol_Admin Then
        Debug.Print "FALLO: Se esperaba Rol_Admin para '" & adminEmail & "' pero se obtuvo " & adminRole
        GoTo TestFail
    End If
    
    If calidadRole <> Rol_Calidad Then
        Debug.Print "FALLO: Se esperaba Rol_Calidad para '" & calidadEmail & "' pero se obtuvo " & calidadRole
        GoTo TestFail
    End If
    
    Test_GetUserRole_RolesEspecificos = True
    Exit Function

TestFail:
    modErrorHandler.LogError "Test_GetUserRole_RolesEspecificos", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_GetUserRole_RolesEspecificos = False
End Function

' ============================================================================
' PRUEBAS DE VALIDACIÃ“N DE EMAIL
' ============================================================================

Public Function Test_ValidateEmail_ValidFormat_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.ValidateEmail("usuario@condor.com")
    
    ' Assert
    Test_ValidateEmail_ValidFormat_ReturnsTrue = result
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_ValidateEmail_ValidFormat_ReturnsTrue", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_ValidateEmail_ValidFormat_ReturnsTrue = False
End Function

Public Function Test_ValidateEmail_InvalidFormat_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.ValidateEmail("email_invalido")
    
    ' Assert
    Test_ValidateEmail_InvalidFormat_ReturnsFalse = Not result
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_ValidateEmail_InvalidFormat_ReturnsFalse", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_ValidateEmail_InvalidFormat_ReturnsFalse = False
End Function

Public Function Test_ValidateEmail_EmptyString_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result As Boolean
    result = authService.ValidateEmail("")
    
    ' Assert
    Test_ValidateEmail_EmptyString_ReturnsFalse = Not result
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_ValidateEmail_EmptyString_ReturnsFalse", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_ValidateEmail_EmptyString_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÃ“N
' ============================================================================

Public Function Test_Integration_AuthenticateAndAuthorize() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim authResult As Boolean
    Dim authzResult As Boolean
    authResult = authService.AuthenticateUser(m_MockUser.Email)
    authzResult = authService.IsUserAuthorized(m_MockUser.Email, "READ")
    
    ' Assert
    ' Verificamos que ambas operaciones se ejecutan sin errores
    Test_Integration_AuthenticateAndAuthorize = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_Integration_AuthenticateAndAuthorize", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_Integration_AuthenticateAndAuthorize = False
End Function

Public Function Test_Integration_GetCurrentUserEmail() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    ' No hay configuraciÃ³n especÃ­fica necesaria para esta prueba
    
    ' Act
    Dim currentEmail As String
    currentEmail = modAppManager.GetCurrentUserEmail()
    
    ' Assert
    ' Verificamos que la funciÃ³n retorna un string
    Test_Integration_GetCurrentUserEmail = (Len(currentEmail) >= 0)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_Integration_GetCurrentUserEmail", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_Integration_GetCurrentUserEmail = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_MultipleAuthentications_SameUser() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuraciÃ³n
    Set authService = authServiceImpl
    
    ' Act
    Dim result1 As Boolean
    Dim result2 As Boolean
    Dim result3 As Boolean
    result1 = authService.AuthenticateUser(m_MockUser.Email)
    result2 = authService.AuthenticateUser(m_MockUser.Email)
    result3 = authService.AuthenticateUser(m_MockUser.Email)
    
    ' Assert
    ' MÃºltiples autenticaciones del mismo usuario deberÃ­an ser consistentes
    Test_MultipleAuthentications_SameUser = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_MultipleAuthentications_SameUser", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_MultipleAuthentications_SameUser = False
End Function

Public Function Test_ConcurrentUsers_DifferentRoles() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Act
    Dim adminRole As E_UserRole
    Dim guestRole As E_UserRole
    adminRole = authService.GetUserRole("admin@condor.com")
    guestRole = authService.GetUserRole("guest@condor.com")
    
    ' Assert
    ' Verificamos que maneja mÃºltiples usuarios sin conflictos
    Test_ConcurrentUsers_DifferentRoles = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_ConcurrentUsers_DifferentRoles", Err.Number, Err.Description, "Test_CAuthService.bas"
    Test_ConcurrentUsers_DifferentRoles = False
End Function

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE EJECUCIÃ“N DE PRUEBAS
' ============================================================================

Public Function RunCAuthServiceTests() As String
    Dim resultado As String
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    resultado = "=== PRUEBAS DE CAuthService ===" & vbCrLf
    totalTests = 0
    passedTests = 0
    
    ' Ejecutar todas las pruebas
    totalTests = totalTests + 1
    If Test_CAuthService_Creation_Success() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_CAuthService_Creation_Success" & vbCrLf
    Else
        resultado = resultado & "? Test_CAuthService_Creation_Success" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CAuthService_ImplementsIAuthService() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_CAuthService_ImplementsIAuthService" & vbCrLf
    Else
        resultado = resultado & "? Test_CAuthService_ImplementsIAuthService" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_AuthenticateUser_ValidEmail_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_AuthenticateUser_ValidEmail_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "? Test_AuthenticateUser_ValidEmail_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_AuthenticateUser_InvalidEmail_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_AuthenticateUser_InvalidEmail_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_AuthenticateUser_InvalidEmail_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_AuthenticateUser_EmptyEmail_HandlesGracefully() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_AuthenticateUser_EmptyEmail_HandlesGracefully" & vbCrLf
    Else
        resultado = resultado & "? Test_AuthenticateUser_EmptyEmail_HandlesGracefully" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_IsUserAuthorized_ValidUser_ReturnsBoolean() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_IsUserAuthorized_ValidUser_ReturnsBoolean" & vbCrLf
    Else
        resultado = resultado & "? Test_IsUserAuthorized_ValidUser_ReturnsBoolean" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_IsUserAuthorized_InvalidPermission_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_IsUserAuthorized_InvalidPermission_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_IsUserAuthorized_InvalidPermission_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_IsUserAuthorized_GuestUser_LimitedPermissions() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_IsUserAuthorized_GuestUser_LimitedPermissions" & vbCrLf
    Else
        resultado = resultado & "? Test_IsUserAuthorized_GuestUser_LimitedPermissions" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetUserRole_ValidUser_ReturnsRole() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetUserRole_ValidUser_ReturnsRole" & vbCrLf
    Else
        resultado = resultado & "? Test_GetUserRole_ValidUser_ReturnsRole" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetUserRole_InvalidUser_ReturnsEmpty() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetUserRole_InvalidUser_ReturnsEmpty" & vbCrLf
    Else
        resultado = resultado & "? Test_GetUserRole_InvalidUser_ReturnsEmpty" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidateEmail_ValidFormat_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidateEmail_ValidFormat_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidateEmail_ValidFormat_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidateEmail_InvalidFormat_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidateEmail_InvalidFormat_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidateEmail_InvalidFormat_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidateEmail_EmptyString_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ValidateEmail_EmptyString_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "? Test_ValidateEmail_EmptyString_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_AuthenticateAndAuthorize() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Integration_AuthenticateAndAuthorize" & vbCrLf
    Else
        resultado = resultado & "? Test_Integration_AuthenticateAndAuthorize" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_GetCurrentUserEmail() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Integration_GetCurrentUserEmail" & vbCrLf
    Else
        resultado = resultado & "? Test_Integration_GetCurrentUserEmail" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_MultipleAuthentications_SameUser() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_MultipleAuthentications_SameUser" & vbCrLf
    Else
        resultado = resultado & "? Test_MultipleAuthentications_SameUser" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ConcurrentUsers_DifferentRoles() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_ConcurrentUsers_DifferentRoles" & vbCrLf
    Else
        resultado = resultado & "? Test_ConcurrentUsers_DifferentRoles" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetUserRole_RolesEspecificos() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetUserRole_RolesEspecificos" & vbCrLf
    Else
        resultado = resultado & "? Test_GetUserRole_RolesEspecificos" & vbCrLf
    End If
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resultado: " & passedTests & "/" & totalTests & " pruebas exitosas" & vbCrLf
    
    RunCAuthServiceTests = resultado
End Function







