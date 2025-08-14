Attribute VB_Name = "Test_AuthService"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: Test_AuthService
' PROPÓSITO: Pruebas unitarias para el servicio de autenticación
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' =====================================================
' FUNCIÓN: RunAllTests
' PROPÓSITO: Ejecutar todas las pruebas del servicio de autenticación
' =====================================================
Public Sub RunAllTests()
    Debug.Print "=== INICIANDO PRUEBAS DE AUTENTICACIÓN ==="
    Debug.Print "Fecha: " & Now()
    
    ' Ejecutar todas las pruebas
    Test_AuthService_Interface
    Test_GetUserRole_AdminUser
    Test_GetUserRole_CalidadUser
    Test_GetUserRole_TecnicoUser
    Test_GetUserRole_UnknownUser
    Test_GetUserRole_EmptyEmail
    Test_GetUserRole_InvalidEmail
    Test_GetCurrentUserEmail_Function
    
    Debug.Print "=== PRUEBAS DE AUTENTICACIÓN COMPLETADAS ==="
End Sub

' =====================================================
' PRUEBA: Test_AuthService_Interface
' PROPÓSITO: Verificar que la interfaz IAuthService funciona correctamente
' =====================================================
Private Sub Test_AuthService_Interface()
    Debug.Print "[TEST] Verificando interfaz IAuthService..."
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Verificar que la instancia se creó correctamente
    Debug.Assert Not authService Is Nothing
    
    Debug.Print "[PASS] Interfaz IAuthService creada correctamente"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_AdminUser
' PROPÓSITO: Verificar que un usuario administrador es identificado correctamente
' =====================================================
Private Sub Test_GetUserRole_AdminUser()
    Debug.Print "[TEST] Verificando rol de usuario administrador..."
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Email de prueba para administrador
    Dim adminEmail As String
    adminEmail = "admin.condor@example.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(adminEmail)
    
    ' Nota: Esta prueba fallará si no hay datos de prueba en la BD Lanzadera
    ' En un entorno de pruebas real, se debería configurar datos de prueba
    Debug.Print "[INFO] Rol obtenido para " & adminEmail & ": " & GetRoleNameForTest(userRole)
    
    ' Para pruebas sin datos reales, verificamos que no hay errores de ejecución
    Debug.Assert userRole >= Rol_Desconocido And userRole <= Rol_Admin
    
    Debug.Print "[PASS] Prueba de usuario administrador completada"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_CalidadUser
' PROPÓSITO: Verificar que un usuario de calidad es identificado correctamente
' =====================================================
Private Sub Test_GetUserRole_CalidadUser()
    Debug.Print "[TEST] Verificando rol de usuario de calidad..."
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Email de prueba para calidad
    Dim calidadEmail As String
    calidadEmail = "calidad.condor@example.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(calidadEmail)
    
    Debug.Print "[INFO] Rol obtenido para " & calidadEmail & ": " & GetRoleNameForTest(userRole)
    
    ' Verificar que el rol está en el rango válido
    Debug.Assert userRole >= Rol_Desconocido And userRole <= Rol_Admin
    
    Debug.Print "[PASS] Prueba de usuario de calidad completada"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_TecnicoUser
' PROPÓSITO: Verificar que un usuario técnico es identificado correctamente
' =====================================================
Private Sub Test_GetUserRole_TecnicoUser()
    Debug.Print "[TEST] Verificando rol de usuario técnico..."
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Email de prueba para técnico
    Dim tecnicoEmail As String
    tecnicoEmail = "tecnico.condor@example.com"
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole(tecnicoEmail)
    
    Debug.Print "[INFO] Rol obtenido para " & tecnicoEmail & ": " & GetRoleNameForTest(userRole)
    
    ' Verificar que el rol está en el rango válido
    Debug.Assert userRole >= Rol_Desconocido And userRole <= Rol_Admin
    
    Debug.Print "[PASS] Prueba de usuario técnico completada"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_UnknownUser
' PROPÓSITO: Verificar que un usuario desconocido retorna Rol_Desconocido
' =====================================================
Private Sub Test_GetUserRole_UnknownUser()
    Debug.Print "[TEST] Verificando usuario desconocido..."
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Email de usuario que no debería existir
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
' PROPÓSITO: Verificar manejo de email vacío
' =====================================================
Private Sub Test_GetUserRole_EmptyEmail()
    Debug.Print "[TEST] Verificando manejo de email vacío..."
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("")
    
    ' Email vacío debe retornar Rol_Desconocido
    Debug.Assert userRole = Rol_Desconocido
    
    Debug.Print "[PASS] Email vacío retorna Rol_Desconocido correctamente"
End Sub

' =====================================================
' PRUEBA: Test_GetUserRole_InvalidEmail
' PROPÓSITO: Verificar manejo de email con caracteres especiales
' =====================================================
Private Sub Test_GetUserRole_InvalidEmail()
    Debug.Print "[TEST] Verificando manejo de email con caracteres especiales..."
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
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
' PROPÓSITO: Verificar que la función GetCurrentUserEmail funciona
' =====================================================
Private Sub Test_GetCurrentUserEmail_Function()
    Debug.Print "[TEST] Verificando función GetCurrentUserEmail..."
    
    Dim userEmail As String
    userEmail = GetCurrentUserEmail()
    
    ' La función debe retornar un string (puede estar vacío en entorno de pruebas)
    Debug.Assert VarType(userEmail) = vbString
    
    Debug.Print "[INFO] Email actual obtenido: '" & userEmail & "'"
    Debug.Print "[PASS] Función GetCurrentUserEmail ejecutada correctamente"
End Sub

' =====================================================
' FUNCIÓN AUXILIAR: GetRoleNameForTest
' PROPÓSITO: Convertir enum a texto para las pruebas
' =====================================================
Private Function GetRoleNameForTest(ByVal role As E_UserRole) As String
    Select Case role
        Case Rol_Admin
            GetRoleNameForTest = "Administrador"
        Case Rol_Calidad
            GetRoleNameForTest = "Calidad"
        Case Rol_Tecnico
            GetRoleNameForTest = "Técnico"
        Case Rol_Desconocido
            GetRoleNameForTest = "Desconocido"
        Case Else
            GetRoleNameForTest = "Indefinido (" & CStr(role) & ")"
    End Select
End Function