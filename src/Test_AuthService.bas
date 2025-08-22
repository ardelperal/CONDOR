Option Compare Database
Option Explicit
' =====================================================
' MODULO: Test_AuthService
' PROPOSITO: Pruebas de integración para CAuthService
' DESCRIPCION: Valida GetUserRole con mocks de base de datos
'              Implementa patrón AAA (Arrange, Act, Assert)
' =====================================================

' --- FUNCIÓN PRINCIPAL DE LA SUITE ---
Public Function Test_AuthService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.SuiteName = "Test_AuthService"

    ' Añadir las dos pruebas de integración
    suiteResult.AddTest "Test_GetUserRole_UsuarioValido_DevuelveRolCorrecto", Test_GetUserRole_UsuarioValido_DevuelveRolCorrecto()
    suiteResult.AddTest "Test_GetUserRole_UsuarioInvalido_DevuelveRolDesconocido", Test_GetUserRole_UsuarioInvalido_DevuelveRolDesconocido()

    Set Test_AuthService_RunAll = suiteResult
End Function

' ==============================
' PRUEBAS DE INTEGRACIÓN
' ==============================

' Prueba: GetUserRole con usuario válido devuelve rol correcto
Public Function Test_GetUserRole_UsuarioValido_DevuelveRolCorrecto() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.TestName = "Test_GetUserRole_UsuarioValido_DevuelveRolCorrecto"
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockSolicitudRepository
    Dim authService As New CAuthService
    
    ' Configurar mock para simular usuario con rol Calidad
    Call mockConfig.SetValue("LANZADERADBPATH", "C:\Test\Lanzadera.accdb")
    Call mockConfig.SetValue("DATABASEPASSWORD", "testpass")
    
    ' Configurar mock recordset para simular datos de usuario de Calidad
    ' TODO: Crear recordset falso que simule EsUsuarioCalidad = "Sí"
    ' mockRepository.SetMockRecordset(recordsetConDatosCalidad)
    
    ' Inicializar AuthService con dependencias mock
    authService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Act
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("usuario.calidad@test.com")
    
    ' Assert
    Call modAssert.AssertEqual(Rol_Calidad, userRole, "GetUserRole debería devolver Rol_Calidad para usuario válido")
    
    testResult.Passed = True
    testResult.Message = "Test completado exitosamente"
    Set Test_GetUserRole_UsuarioValido_DevuelveRolCorrecto = testResult
    
    Exit Function
    
ErrorHandler:
    testResult.Passed = False
    testResult.Message = "Error en test: " & Err.Description
    Set Test_GetUserRole_UsuarioValido_DevuelveRolCorrecto = testResult
End Function

' Prueba: GetUserRole con usuario inválido devuelve rol desconocido
Public Function Test_GetUserRole_UsuarioInvalido_DevuelveRolDesconocido() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As New CTestResult
    testResult.TestName = "Test_GetUserRole_UsuarioInvalido_DevuelveRolDesconocido"
    
    ' Arrange
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockSolicitudRepository
    Dim authService As New CAuthService
    
    ' Configurar mock para simular base de datos sin datos
    Call mockConfig.SetValue("LANZADERADBPATH", "C:\Test\Lanzadera.accdb")
    Call mockConfig.SetValue("DATABASEPASSWORD", "testpass")
    
    ' Configurar mock recordset para simular recordset vacío (EOF = True)
    ' TODO: Crear recordset falso vacío
    ' mockRepository.SetMockRecordset(recordsetVacio)
    
    ' Inicializar AuthService con dependencias mock
    authService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Act
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("usuario.inexistente@test.com")
    
    ' Assert
    Call modAssert.AssertEqual(Rol_Desconocido, userRole, "GetUserRole debería devolver Rol_Desconocido para usuario inválido")
    
    testResult.Passed = True
    testResult.Message = "Test completado exitosamente"
    Set Test_GetUserRole_UsuarioInvalido_DevuelveRolDesconocido = testResult
    
    Exit Function
    
ErrorHandler:
    testResult.Passed = False
    testResult.Message = "Error en test: " & Err.Description
    Set Test_GetUserRole_UsuarioInvalido_DevuelveRolDesconocido = testResult
End Function











