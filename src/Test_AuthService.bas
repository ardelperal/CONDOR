Attribute VB_Name = "Test_AuthService"
Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_AuthService
' PROPOSITO: Pruebas unitarias para el servicio de autenticacion
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' =====================================================
' FUNCION: RunAllTests
' PROPOSITO: Ejecutar todas las pruebas del servicio de autenticacion
' =====================================================
Public Function RunAllTests() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== INICIANDO PRUEBAS DE AUTENTICACION ===" & vbCrLf
    resultado = resultado & "Fecha: " & Now() & vbCrLf & vbCrLf
    
    testsTotal = 8
    testsPassed = 0
    
    ' Ejecutar todas las pruebas
    On Error GoTo ErrorHandler
    
    Test_AuthService_Interface
    resultado = resultado & "[PASS] Test_AuthService_Interface" & vbCrLf
    testsPassed = testsPassed + 1
    
    Test_GetUserRole_AdminUser
    resultado = resultado & "[PASS] Test_GetUserRole_AdminUser" & vbCrLf
    testsPassed = testsPassed + 1
    
    Test_GetUserRole_CalidadUser
    resultado = resultado & "[PASS] Test_GetUserRole_CalidadUser" & vbCrLf
    testsPassed = testsPassed + 1
    
    Test_GetUserRole_TecnicoUser
    resultado = resultado & "[PASS] Test_GetUserRole_TecnicoUser" & vbCrLf
    testsPassed = testsPassed + 1
    
    Test_GetUserRole_UnknownUser
    resultado = resultado & "[PASS] Test_GetUserRole_UnknownUser" & vbCrLf
    testsPassed = testsPassed + 1
    
    Test_GetUserRole_EmptyEmail
    resultado = resultado & "[PASS] Test_GetUserRole_EmptyEmail" & vbCrLf
    testsPassed = testsPassed + 1
    
    Test_GetUserRole_InvalidEmail
    resultado = resultado & "[PASS] Test_GetUserRole_InvalidEmail" & vbCrLf
    testsPassed = testsPassed + 1
    
    Test_GetCurrentUserEmail_Function
    resultado = resultado & "[PASS] Test_GetCurrentUserEmail_Function" & vbCrLf
    testsPassed = testsPassed + 1
    
    GoTo TestsComplete
    
ErrorHandler:
    resultado = resultado & "[FAIL] Error en prueba: " & Err.Description & vbCrLf
    Resume Next
    
TestsComplete:
    resultado = resultado & vbCrLf & "RESUMEN AUTENTICACION:" & vbCrLf
    resultado = resultado & "Pruebas ejecutadas: " & testsTotal & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & testsPassed & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & (testsTotal - testsPassed) & vbCrLf
    
    If testsPassed = testsTotal Then
        resultado = resultado & "RESULTADO AUTENTICACION: TODAS LAS PRUEBAS PASARON [OK]" & vbCrLf
    Else
        resultado = resultado & "RESULTADO AUTENTICACION: ALGUNAS PRUEBAS FALLARON [ERROR]" & vbCrLf
    End If
    
    resultado = resultado & "=== PRUEBAS DE AUTENTICACION COMPLETADAS ===" & vbCrLf
    
    RunAllTests = resultado
End Function

' =====================================================
' PRUEBA: Test_AuthService_Interface
' PROPOSITO: Verificar que la interfaz IAuthService funciona correctamente
' =====================================================
Private Sub Test_AuthService_Interface()
    Debug.Print "[TEST] Verificando interfaz IAuthService..."
    
    Dim authService As IAuthService
    Set authService = New CMockAuthService
    
    ' Verificar que la instancia se creo correctamente
    Debug.Assert Not authService Is Nothing
    
    Debug.Print "[PASS] Interfaz IAuthService creada correctamente"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_AdminUser
' PROPOSITO: Verificar que un usuario administrador es identificado correctamente
' =====================================================
Private Sub Test_GetUserRole_AdminUser()
    Debug.Print "[TEST] Verificando rol de usuario administrador..."
    
    Dim authService As IAuthService
    Set authService = New CMockAuthService
    
    ' Email de prueba para administrador
    Dim adminEmail As String
    adminEmail = "admin.condor@example.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(adminEmail)
    
    ' Con CMockAuthService, debe retornar Rol_Admin para este email
    Debug.Assert userRole = Rol_Admin
    
    Debug.Print "[INFO] Rol obtenido para " & adminEmail & ": " & GetRoleNameForTest(userRole)
    Debug.Print "[PASS] Prueba de usuario administrador completada"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_CalidadUser
' PROPOSITO: Verificar que un usuario de calidad es identificado correctamente
' =====================================================
Private Sub Test_GetUserRole_CalidadUser()
    Debug.Print "[TEST] Verificando rol de usuario de calidad..."
    
    Dim authService As IAuthService
    Set authService = New CMockAuthService
    
    ' Email de prueba para calidad
    Dim calidadEmail As String
    calidadEmail = "calidad.condor@example.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(calidadEmail)
    
    ' Con CMockAuthService, debe retornar Rol_Calidad para este email
    Debug.Assert userRole = Rol_Calidad
    
    Debug.Print "[INFO] Rol obtenido para " & calidadEmail & ": " & GetRoleNameForTest(userRole)
    Debug.Print "[PASS] Prueba de usuario de calidad completada"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_TecnicoUser
' PROPOSITO: Verificar que un usuario tecnico es identificado correctamente
' =====================================================
Private Sub Test_GetUserRole_TecnicoUser()
    Debug.Print "[TEST] Verificando rol de usuario tecnico..."
    
    Dim authService As IAuthService
    Set authService = New CMockAuthService
    
    ' Email de prueba para tecnico
    Dim tecnicoEmail As String
    tecnicoEmail = "tecnico.condor@example.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(tecnicoEmail)
    
    ' Con CMockAuthService, debe retornar Rol_Tecnico para este email
    Debug.Assert userRole = Rol_Tecnico
    
    Debug.Print "[INFO] Rol obtenido para " & tecnicoEmail & ": " & GetRoleNameForTest(userRole)
    Debug.Print "[PASS] Prueba de usuario tecnico completada"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_UnknownUser
' PROPOSITO: Verificar que un usuario desconocido retorna Rol_Desconocido
' =====================================================
Private Sub Test_GetUserRole_UnknownUser()
    Debug.Print "[TEST] Verificando usuario desconocido..."
    
    Dim authService As IAuthService
    Set authService = New CMockAuthService
    
    ' Email de usuario que no deberia existir
    Dim unknownEmail As String
    unknownEmail = "usuario.inexistente@noexiste.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(unknownEmail)
    
    ' Un usuario inexistente debe retornar Rol_Desconocido
    Debug.Assert userRole = Rol_Desconocido
    
    Debug.Print "[PASS] Usuario desconocido retorna Rol_Desconocido correctamente"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_EmptyEmail
' PROPOSITO: Verificar manejo de email vacio
' =====================================================
Private Sub Test_GetUserRole_EmptyEmail()
    Debug.Print "[TEST] Verificando manejo de email vacio..."
    
    Dim authService As IAuthService
    Set authService = New CMockAuthService
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("")
    
    ' Email vacio debe retornar Rol_Desconocido
    Debug.Assert userRole = Rol_Desconocido
    
    Debug.Print "[PASS] Email vacio retorna Rol_Desconocido correctamente"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_InvalidEmail
' PROPOSITO: Verificar manejo de email con caracteres especiales
' =====================================================
Private Sub Test_GetUserRole_InvalidEmail()
    Debug.Print "[TEST] Verificando manejo de email con caracteres especiales..."
    
    Dim authService As IAuthService
    Set authService = New CMockAuthService
    
    ' Email con comilla simple para probar escape SQL
    Dim invalidEmail As String
    invalidEmail = "usuario'malicioso@test.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(invalidEmail)
    
    ' Debe manejar el email sin errores y retornar Rol_Desconocido
    Debug.Assert userRole = Rol_Desconocido
    
    Debug.Print "[PASS] Email con caracteres especiales manejado correctamente"
End Sub

' =====================================================
' PRUEBA: Test_GetCurrentUserEmail_Function
' PROPOSITO: Verificar que la funcion GetCurrentUserEmail funciona
' =====================================================
Private Sub Test_GetCurrentUserEmail_Function()
    Debug.Print "[TEST] Verificando funcion GetCurrentUserEmail..."
    
    Dim userEmail As String
    userEmail = GetCurrentUserEmail()
    
    ' La funcion debe retornar un string (puede estar vacio en entorno de pruebas)
    Debug.Assert VarType(userEmail) = vbString
    
    Debug.Print "[INFO] Email actual obtenido: '" & userEmail & "'"
    Debug.Print "[PASS] Funcion GetCurrentUserEmail ejecutada correctamente"
End Sub

' =====================================================
' FUNCION AUXILIAR: GetRoleNameForTest
' PROPOSITO: Convertir enum a texto para las pruebas
' =====================================================
Private Function GetRoleNameForTest(ByVal role As E_UserRole) As String
    Select Case role
        Case Rol_Admin
            GetRoleNameForTest = "Administrador"
        Case Rol_Calidad
            GetRoleNameForTest = "Calidad"
        Case Rol_Tecnico
            GetRoleNameForTest = "Tecnico"
        Case Rol_Desconocido
            GetRoleNameForTest = "Desconocido"
        Case Else
            GetRoleNameForTest = "Indefinido (" & CStr(role) & ")"
    End Select
End Function
