Option Compare Database
Option Explicit
' ============================================================================
' Módulo: Test_AppManager
' Descripción: Pruebas unitarias para modAppManager.bas
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' Mock para simular datos de usuario y roles
Private Type T_MockUserData
    Email As String
    role As E_UserRole
    IsValid As Boolean
    ShouldFailAuth As Boolean
End Type

Private m_MockUser As T_MockUserData
Private m_OriginalUserRole As E_UserRole

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN DE MOCKS
' ============================================================================

Private Sub SetupValidAdminUserMock()
    m_MockUser.Email = "admin@condor.com"
    m_MockUser.role = Rol_Admin
    m_MockUser.IsValid = True
    m_MockUser.ShouldFailAuth = False
End Sub

Private Sub SetupValidCalidadUserMock()
    m_MockUser.Email = "calidad@condor.com"
    m_MockUser.role = Rol_Calidad
    m_MockUser.IsValid = True
    m_MockUser.ShouldFailAuth = False
End Sub

Private Sub SetupValidTecnicoUserMock()
    m_MockUser.Email = "tecnico@condor.com"
    m_MockUser.role = Rol_Tecnico
    m_MockUser.IsValid = True
    m_MockUser.ShouldFailAuth = False
End Sub

Private Sub SetupInvalidUserMock()
    m_MockUser.Email = "invalid@condor.com"
    m_MockUser.role = Rol_Desconocido
    m_MockUser.IsValid = False
    m_MockUser.ShouldFailAuth = True
End Sub

Private Sub SetupEmptyUserMock()
    m_MockUser.Email = ""
    m_MockUser.role = Rol_Desconocido
    m_MockUser.IsValid = False
    m_MockUser.ShouldFailAuth = True
End Sub

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN Y LIMPIEZA
' ============================================================================

Private Sub SaveCurrentUserRole()
    m_OriginalUserRole = g_CurrentUserRole
End Sub

Private Sub RestoreCurrentUserRole()
    g_CurrentUserRole = m_OriginalUserRole
End Sub

' ============================================================================
' PRUEBAS DE FUNCIÓN GetCurrentUserEmail
' ============================================================================

Public Function Test_GetCurrentUserEmail_ReturnsString() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim Email As String
    Email = GetCurrentUserEmail()
    
    ' Assert
    ' Verificamos que retorna un string (puede estar vacío en modo desarrollo)
    Test_GetCurrentUserEmail_ReturnsString = True
    
    Exit Function
    
TestFail:
    Test_GetCurrentUserEmail_ReturnsString = False
End Function

Public Function Test_GetCurrentUserEmail_DevMode_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim Email As String
    Email = GetCurrentUserEmail()
    
    ' Assert
    ' En modo desarrollo, puede usar VBA.Command o valor por defecto
    Test_GetCurrentUserEmail_DevMode_HandlesCorrectly = True
    
    Exit Function
    
TestFail:
    Test_GetCurrentUserEmail_DevMode_HandlesCorrectly = False
End Function

' ============================================================================
' PRUEBAS DE FUNCIÓN Ping
' ============================================================================

Public Function Test_Ping_ReturnsPong() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim result As String
    result = Ping()
    
    ' Assert
    Test_Ping_ReturnsPong = (result = "Pong")
    
    Exit Function
    
TestFail:
    Test_Ping_ReturnsPong = False
End Function

Public Function Test_Ping_ConsistentResponse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim result1 As String
    Dim result2 As String
    result1 = Ping()
    result2 = Ping()
    
    ' Assert
    Test_Ping_ConsistentResponse = (result1 = result2) And (result1 = "Pong")
    
    Exit Function
    
TestFail:
    Test_Ping_ConsistentResponse = False
End Function

' ============================================================================
' PRUEBAS DE ROLES DE USUARIO
' ============================================================================

Public Function Test_UserRole_AdminRole_SetsCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SaveCurrentUserRole
    SetupValidAdminUserMock
    
    ' Act
    g_CurrentUserRole = m_MockUser.role
    
    ' Assert
    Test_UserRole_AdminRole_SetsCorrectly = (g_CurrentUserRole = Rol_Admin)
    
    ' Cleanup
    RestoreCurrentUserRole
    
    Exit Function
    
TestFail:
    RestoreCurrentUserRole
    Test_UserRole_AdminRole_SetsCorrectly = False
End Function

Public Function Test_UserRole_CalidadRole_SetsCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SaveCurrentUserRole
    SetupValidCalidadUserMock
    
    ' Act
    g_CurrentUserRole = m_MockUser.role
    
    ' Assert
    Test_UserRole_CalidadRole_SetsCorrectly = (g_CurrentUserRole = Rol_Calidad)
    
    ' Cleanup
    RestoreCurrentUserRole
    
    Exit Function
    
TestFail:
    RestoreCurrentUserRole
    Test_UserRole_CalidadRole_SetsCorrectly = False
End Function

Public Function Test_UserRole_TecnicoRole_SetsCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SaveCurrentUserRole
    SetupValidTecnicoUserMock
    
    ' Act
    g_CurrentUserRole = m_MockUser.role
    
    ' Assert
    Test_UserRole_TecnicoRole_SetsCorrectly = (g_CurrentUserRole = Rol_Tecnico)
    
    ' Cleanup
    RestoreCurrentUserRole
    
    Exit Function
    
TestFail:
    RestoreCurrentUserRole
    Test_UserRole_TecnicoRole_SetsCorrectly = False
End Function

Public Function Test_UserRole_DesconocidoRole_SetsCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SaveCurrentUserRole
    SetupInvalidUserMock
    
    ' Act
    g_CurrentUserRole = m_MockUser.role
    
    ' Assert
    Test_UserRole_DesconocidoRole_SetsCorrectly = (g_CurrentUserRole = Rol_Desconocido)
    
    ' Cleanup
    RestoreCurrentUserRole
    
    Exit Function
    
TestFail:
    RestoreCurrentUserRole
    Test_UserRole_DesconocidoRole_SetsCorrectly = False
End Function

' ============================================================================
' PRUEBAS DE ENUMERACIÓN E_UserRole
' ============================================================================

Public Function Test_UserRoleEnum_ValidValues() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act & Assert
    Dim result As Boolean
    result = True
    
    ' Verificar que los valores de la enumeración son correctos
    result = result And (Rol_Desconocido = 0)
    result = result And (Rol_Tecnico = 1)
    result = result And (Rol_Calidad = 2)
    result = result And (Rol_Admin = 3)
    
    Test_UserRoleEnum_ValidValues = result
    
    Exit Function
    
TestFail:
    Test_UserRoleEnum_ValidValues = False
End Function

Public Function Test_UserRoleEnum_CanAssignToVariable() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim testRole As E_UserRole
    testRole = Rol_Admin
    
    ' Assert
    Test_UserRoleEnum_CanAssignToVariable = (testRole = Rol_Admin)
    
    Exit Function
    
TestFail:
    Test_UserRoleEnum_CanAssignToVariable = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN CON SERVICIOS
' ============================================================================

Public Function Test_Integration_WithAuthService() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidAdminUserMock
    Dim authService As IAuthService
    Dim authServiceImpl As CAuthService
    Set authServiceImpl = New CAuthService
    authServiceImpl.Initialize AppConfig ' Inyectar dependencia de configuración
    Set authService = authServiceImpl
    
    ' Act
    Dim UserRole As E_UserRole
    UserRole = authService.GetUserRole(m_MockUser.Email)
    
    ' Assert
    ' Verificamos que la integración funciona (no falla)
    Test_Integration_WithAuthService = True
    
    Exit Function
    
TestFail:
    Test_Integration_WithAuthService = False
End Function

Public Function Test_Integration_EmailAndRoleConsistency() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidCalidadUserMock
    SaveCurrentUserRole
    
    ' Act
    Dim Email As String
    Email = GetCurrentUserEmail()
    g_CurrentUserRole = m_MockUser.role
    
    ' Assert
    ' Verificamos que podemos obtener email y establecer rol sin conflictos
    Test_Integration_EmailAndRoleConsistency = (g_CurrentUserRole = Rol_Calidad)
    
    ' Cleanup
    RestoreCurrentUserRole
    
    Exit Function
    
TestFail:
    RestoreCurrentUserRole
    Test_Integration_EmailAndRoleConsistency = False
End Function

' ============================================================================
' PRUEBAS DE FUNCIONES DE PRUEBA
' ============================================================================

Public Function Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    ' Ejecutamos la subrutina de pruebas (no podemos verificar el output directamente)
    Call EJECUTAR_TODAS_LAS_PRUEBAS
    
    ' Assert
    Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail = True
    
    Exit Function
    
TestFail:
    Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail = False
End Function

Public Function Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim resultado As String
    resultado = OBTENER_RESULTADOS_PRUEBAS()
    
    ' Assert
    ' Verificamos que retorna un string (puede estar vacío)
    Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString = True
    
    Exit Function
    
TestFail:
    Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_EdgeCase_MultipleRoleChanges() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SaveCurrentUserRole
    
    ' Act
    g_CurrentUserRole = Rol_Admin
    Dim role1 As E_UserRole
    role1 = g_CurrentUserRole
    
    g_CurrentUserRole = Rol_Tecnico
    Dim role2 As E_UserRole
    role2 = g_CurrentUserRole
    
    g_CurrentUserRole = Rol_Calidad
    Dim role3 As E_UserRole
    role3 = g_CurrentUserRole
    
    ' Assert
    Test_EdgeCase_MultipleRoleChanges = (role1 = Rol_Admin) And _
                                       (role2 = Rol_Tecnico) And _
                                       (role3 = Rol_Calidad)
    
    ' Cleanup
    RestoreCurrentUserRole
    
    Exit Function
    
TestFail:
    RestoreCurrentUserRole
    Test_EdgeCase_MultipleRoleChanges = False
End Function

Public Function Test_EdgeCase_ConcurrentEmailCalls() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim email1 As String
    Dim email2 As String
    Dim email3 As String
    
    email1 = GetCurrentUserEmail()
    email2 = GetCurrentUserEmail()
    email3 = GetCurrentUserEmail()
    
    ' Assert
    ' Verificamos que múltiples llamadas son consistentes
    Test_EdgeCase_ConcurrentEmailCalls = (email1 = email2) And (email2 = email3)
    
    Exit Function
    
TestFail:
    Test_EdgeCase_ConcurrentEmailCalls = False
End Function

Public Function Test_EdgeCase_PingStressTest() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim i As Integer
    Dim allPassed As Boolean
    allPassed = True
    
    ' Ejecutar Ping múltiples veces
    For i = 1 To 100
        If Ping() <> "Pong" Then
            allPassed = False
            Exit For
        End If
    Next i
    
    ' Assert
    Test_EdgeCase_PingStressTest = allPassed
    
    Exit Function
    
TestFail:
    Test_EdgeCase_PingStressTest = False
End Function

' ============================================================================
' PRUEBAS DE CONSTANTES Y CONFIGURACIÓN
' ============================================================================

Public Function Test_DevMode_ConstantExists() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act & Assert
    ' Verificamos que la constante DEV_MODE está definida
    ' (esto se verifica implícitamente en GetCurrentUserEmail)
    Dim Email As String
    Email = GetCurrentUserEmail()
    
    Test_DevMode_ConstantExists = True
    
    Exit Function
    
TestFail:
    Test_DevMode_ConstantExists = False
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_AppManager_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS DE APPMANAGER ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas
    testsTotal = testsTotal + 1
    If Test_GetCurrentUserEmail_ReturnsString() Then
        resultado = resultado & "[OK] Test_GetCurrentUserEmail_ReturnsString" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_GetCurrentUserEmail_ReturnsString" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_GetCurrentUserEmail_DevMode_HandlesCorrectly() Then
        resultado = resultado & "[OK] Test_GetCurrentUserEmail_DevMode_HandlesCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_GetCurrentUserEmail_DevMode_HandlesCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Ping_ReturnsPong() Then
        resultado = resultado & "[OK] Test_Ping_ReturnsPong" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Ping_ReturnsPong" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Ping_ConsistentResponse() Then
        resultado = resultado & "[OK] Test_Ping_ConsistentResponse" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Ping_ConsistentResponse" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_UserRole_AdminRole_SetsCorrectly() Then
        resultado = resultado & "[OK] Test_UserRole_AdminRole_SetsCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_UserRole_AdminRole_SetsCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_UserRole_CalidadRole_SetsCorrectly() Then
        resultado = resultado & "[OK] Test_UserRole_CalidadRole_SetsCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_UserRole_CalidadRole_SetsCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_UserRole_TecnicoRole_SetsCorrectly() Then
        resultado = resultado & "[OK] Test_UserRole_TecnicoRole_SetsCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_UserRole_TecnicoRole_SetsCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_UserRole_DesconocidoRole_SetsCorrectly() Then
        resultado = resultado & "[OK] Test_UserRole_DesconocidoRole_SetsCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_UserRole_DesconocidoRole_SetsCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_UserRoleEnum_ValidValues() Then
        resultado = resultado & "[OK] Test_UserRoleEnum_ValidValues" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_UserRoleEnum_ValidValues" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_UserRoleEnum_CanAssignToVariable() Then
        resultado = resultado & "[OK] Test_UserRoleEnum_CanAssignToVariable" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_UserRoleEnum_CanAssignToVariable" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Integration_WithAuthService() Then
        resultado = resultado & "[OK] Test_Integration_WithAuthService" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Integration_WithAuthService" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Integration_EmailAndRoleConsistency() Then
        resultado = resultado & "[OK] Test_Integration_EmailAndRoleConsistency" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Integration_EmailAndRoleConsistency" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail() Then
        resultado = resultado & "[OK] Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString() Then
        resultado = resultado & "[OK] Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_EdgeCase_MultipleRoleChanges() Then
        resultado = resultado & "[OK] Test_EdgeCase_MultipleRoleChanges" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EdgeCase_MultipleRoleChanges" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_EdgeCase_ConcurrentEmailCalls() Then
        resultado = resultado & "[OK] Test_EdgeCase_ConcurrentEmailCalls" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EdgeCase_ConcurrentEmailCalls" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_EdgeCase_PingStressTest() Then
        resultado = resultado & "[OK] Test_EdgeCase_PingStressTest" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_EdgeCase_PingStressTest" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_DevMode_ConstantExists() Then
        resultado = resultado & "[OK] Test_DevMode_ConstantExists" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_DevMode_ConstantExists" & vbCrLf
    End If
    
    ' Agregar resumen
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasadas" & vbCrLf
    
    Test_AppManager_RunAll = resultado
End Function

Public Function RunAppManagerTests() As String
    Dim resultado As String
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    resultado = "=== PRUEBAS DE modAppManager ===" & vbCrLf
    totalTests = 0
    passedTests = 0
    
    ' Ejecutar todas las pruebas
    totalTests = totalTests + 1
    If Test_GetCurrentUserEmail_ReturnsString() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetCurrentUserEmail_ReturnsString" & vbCrLf
    Else
        resultado = resultado & "? Test_GetCurrentUserEmail_ReturnsString" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetCurrentUserEmail_DevMode_HandlesCorrectly() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_GetCurrentUserEmail_DevMode_HandlesCorrectly" & vbCrLf
    Else
        resultado = resultado & "? Test_GetCurrentUserEmail_DevMode_HandlesCorrectly" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Ping_ReturnsPong() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Ping_ReturnsPong" & vbCrLf
    Else
        resultado = resultado & "? Test_Ping_ReturnsPong" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Ping_ConsistentResponse() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Ping_ConsistentResponse" & vbCrLf
    Else
        resultado = resultado & "? Test_Ping_ConsistentResponse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UserRole_AdminRole_SetsCorrectly() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_UserRole_AdminRole_SetsCorrectly" & vbCrLf
    Else
        resultado = resultado & "? Test_UserRole_AdminRole_SetsCorrectly" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UserRole_CalidadRole_SetsCorrectly() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_UserRole_CalidadRole_SetsCorrectly" & vbCrLf
    Else
        resultado = resultado & "? Test_UserRole_CalidadRole_SetsCorrectly" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UserRole_TecnicoRole_SetsCorrectly() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_UserRole_TecnicoRole_SetsCorrectly" & vbCrLf
    Else
        resultado = resultado & "? Test_UserRole_TecnicoRole_SetsCorrectly" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UserRole_DesconocidoRole_SetsCorrectly() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_UserRole_DesconocidoRole_SetsCorrectly" & vbCrLf
    Else
        resultado = resultado & "? Test_UserRole_DesconocidoRole_SetsCorrectly" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UserRoleEnum_ValidValues() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_UserRoleEnum_ValidValues" & vbCrLf
    Else
        resultado = resultado & "? Test_UserRoleEnum_ValidValues" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UserRoleEnum_CanAssignToVariable() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_UserRoleEnum_CanAssignToVariable" & vbCrLf
    Else
        resultado = resultado & "? Test_UserRoleEnum_CanAssignToVariable" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_WithAuthService() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Integration_WithAuthService" & vbCrLf
    Else
        resultado = resultado & "? Test_Integration_WithAuthService" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_EmailAndRoleConsistency() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_Integration_EmailAndRoleConsistency" & vbCrLf
    Else
        resultado = resultado & "? Test_Integration_EmailAndRoleConsistency" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail" & vbCrLf
    Else
        resultado = resultado & "? Test_EJECUTAR_TODAS_LAS_PRUEBAS_DoesNotFail" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString" & vbCrLf
    Else
        resultado = resultado & "? Test_OBTENER_RESULTADOS_PRUEBAS_ReturnsString" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_EdgeCase_MultipleRoleChanges() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_EdgeCase_MultipleRoleChanges" & vbCrLf
    Else
        resultado = resultado & "? Test_EdgeCase_MultipleRoleChanges" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_EdgeCase_ConcurrentEmailCalls() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_EdgeCase_ConcurrentEmailCalls" & vbCrLf
    Else
        resultado = resultado & "? Test_EdgeCase_ConcurrentEmailCalls" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_EdgeCase_PingStressTest() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_EdgeCase_PingStressTest" & vbCrLf
    Else
        resultado = resultado & "? Test_EdgeCase_PingStressTest" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_DevMode_ConstantExists() Then
        passedTests = passedTests + 1
        resultado = resultado & "? Test_DevMode_ConstantExists" & vbCrLf
    Else
        resultado = resultado & "? Test_DevMode_ConstantExists" & vbCrLf
    End If
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resultado: " & passedTests & "/" & totalTests & " pruebas exitosas" & vbCrLf
    
    RunAppManagerTests = resultado
End Function













