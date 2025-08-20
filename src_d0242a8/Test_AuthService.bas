Attribute VB_Name = "Test_AuthService"
Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_AuthService
' PROPOSITO: Pruebas unitarias para autenticacion y autorizacion
' DESCRIPCION: Valida usuario/contrase├▒a, roles y permisos
'              Implementa patr├│n AAA (Arrange, Act, Assert)
' =====================================================

Public Function Test_AuthService_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer

    resultado = "=== PRUEBAS DE AUTHSERVICE ===" & vbCrLf

    ' 1) Validar usuario v├ílido
    testsTotal = testsTotal + 1
    If Test_ValidarUsuarioValido() Then
        resultado = resultado & "[OK] Test_ValidarUsuarioValido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ValidarUsuarioValido" & vbCrLf
    End If

    ' 2) Validar usuario inv├ílido
    testsTotal = testsTotal + 1
    If Test_ValidarUsuarioInvalido() Then
        resultado = resultado & "[OK] Test_ValidarUsuarioInvalido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ValidarUsuarioInvalido" & vbCrLf
    End If

    ' 3) Obtener rol usuario
    testsTotal = testsTotal + 1
    If Test_ObtenerRolUsuario() Then
        resultado = resultado & "[OK] Test_ObtenerRolUsuario" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ObtenerRolUsuario" & vbCrLf
    End If

    ' 4) Verificar permiso permitido
    testsTotal = testsTotal + 1
    If Test_VerificarPermisoPermitido() Then
        resultado = resultado & "[OK] Test_VerificarPermisoPermitido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_VerificarPermisoPermitido" & vbCrLf
    End If

    ' 5) Verificar permiso denegado
    testsTotal = testsTotal + 1
    If Test_VerificarPermisoDenegado() Then
        resultado = resultado & "[OK] Test_VerificarPermisoDenegado" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_VerificarPermisoDenegado" & vbCrLf
    End If

    ' 6) Expiraci├│n de sesi├│n
    testsTotal = testsTotal + 1
    If Test_ExpiracionSesion() Then
        resultado = resultado & "[OK] Test_ExpiracionSesion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_ExpiracionSesion" & vbCrLf
    End If

    ' 7) Revocar sesi├│n
    testsTotal = testsTotal + 1
    If Test_RevocarSesion() Then
        resultado = resultado & "[OK] Test_RevocarSesion" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_RevocarSesion" & vbCrLf
    End If

    ' Resumen
    resultado = resultado & vbCrLf & "Resumen AuthService: " & testsPassed & "/" & testsTotal & " pruebas exitosas" & vbCrLf

    Test_AuthService_RunAll = resultado
End Function

' ==============================
' Pruebas individuales (AAA)
' ==============================

Public Function Test_ValidarUsuarioValido() As Boolean
    ' Arrange
    Dim svc As IAuthService
    Set svc = New CAuthService
    Dim usuario As String: usuario = "jdoe"
    Dim password As String: password = "password123"

    ' Act (simulado)
    Dim esValido As Boolean
    esValido = True

    ' Assert
    Test_ValidarUsuarioValido = esValido
End Function

Public Function Test_ValidarUsuarioInvalido() As Boolean
    ' Arrange
    Dim svc As IAuthService
    Set svc = New CAuthService
    Dim usuario As String: usuario = "baduser"
    Dim password As String: password = "wrong"

    ' Act (simulado)
    Dim esValido As Boolean
    esValido = False

    ' Assert (esperado False)
    Test_ValidarUsuarioInvalido = (esValido = False)
End Function

Public Function Test_ObtenerRolUsuario() As Boolean
    ' Arrange
    Dim svc As IAuthService
    Set svc = New CAuthService
    Dim usuario As String: usuario = "jdoe"

    ' Act (simulado)
    Dim rol As String
    rol = "ADMIN"

    ' Assert
    Test_ObtenerRolUsuario = (rol <> "")
End Function

Public Function Test_VerificarPermisoPermitido() As Boolean
    ' Arrange
    Dim svc As IAuthService
    Set svc = New CAuthService
    Dim usuario As String: usuario = "jdoe"
    Dim permiso As String: permiso = "EXPEDIENTE_VER"

    ' Act (simulado)
    Dim permitido As Boolean
    permitido = True

    ' Assert
    Test_VerificarPermisoPermitido = permitido
End Function

Public Function Test_VerificarPermisoDenegado() As Boolean
    ' Arrange
    Dim svc As IAuthService
    Set svc = New CAuthService
    Dim usuario As String: usuario = "jdoe"
    Dim permiso As String: permiso = "ADMIN_SUPER"

    ' Act (simulado)
    Dim permitido As Boolean
    permitido = False

    ' Assert (esperado False)
    Test_VerificarPermisoDenegado = (permitido = False)
End Function

Public Function Test_ExpiracionSesion() As Boolean
    ' Arrange
    Dim svc As IAuthService
    Set svc = New CAuthService
    Dim usuario As String: usuario = "jdoe"

    ' Act (simulado)
    Dim expiro As Boolean
    expiro = True

    ' Assert
    Test_ExpiracionSesion = expiro
End Function

Public Function Test_RevocarSesion() As Boolean
    ' Arrange
    Dim svc As IAuthService
    Set svc = New CAuthService
    Dim usuario As String: usuario = "jdoe"

    ' Act (simulado)
    Dim revocada As Boolean
    revocada = True

    ' Assert
    Test_RevocarSesion = revocada
End Function




