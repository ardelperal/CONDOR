Attribute VB_Name = "Test_AuthService"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' MÃ“DULO DE PRUEBAS UNITARIAS PARA CAuthService
' ============================================================================
' Este mÃ³dulo contiene pruebas unitarias aisladas para CAuthService
' utilizando mocks para todas las dependencias externas.
' Sigue la LecciÃ³n 10: El Aislamiento de las Pruebas Unitarias con Mocks no es Negociable

' FunciÃ³n principal que ejecuta todas las pruebas del mÃ³dulo
Public Function Test_AuthService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_AuthService"
    
    ' Ejecutar todas las pruebas unitarias
    suiteResult.AddTestResult Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador()
    suiteResult.AddTestResult Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad()
    suiteResult.AddTestResult Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico()
    suiteResult.AddTestResult Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido()
    
    Set Test_AuthService_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetUserRole
' ============================================================================

' Prueba que GetUserRole devuelve Rol_Administrador para usuario administrador
Private Function Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockSolicitudRepository
    
    ' Nota: La nueva implementaciÃ³n de IConfig no permite SetValue
    ' Las pruebas deben usar la configuraciÃ³n real del sistema
    
    ' Crear recordset mock para usuario administrador
    Dim mockRecordset As Object
    Set mockRecordset = CreateMockUserRecordset("SÃ­", "No", "No")
    mockRepository.SetMockRecordset mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize configService, mockLogger, mockRepository
    
    ' Act - Ejecutar el mÃ©todo bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("admin@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals Rol_Administrador, userRole, "GetUserRole debe devolver Rol_Administrador para usuario administrador"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
    Set Test_GetUserRole_UsuarioAdministrador_DevuelveRolAdministrador = testResult
End Function

' Prueba que GetUserRole devuelve Rol_Calidad para usuario de calidad
Private Function Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockSolicitudRepository
    
    ' Nota: La nueva implementaciÃ³n de IConfig no permite SetValue
    ' Las pruebas deben usar la configuraciÃ³n real del sistema
    
    ' Crear recordset mock para usuario de calidad
    Dim mockRecordset As Object
    Set mockRecordset = CreateMockUserRecordset("No", "SÃ­", "No")
    mockRepository.SetMockRecordset mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize configService, mockLogger, mockRepository
    
    ' Act - Ejecutar el mÃ©todo bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("calidad@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals Rol_Calidad, userRole, "GetUserRole debe devolver Rol_Calidad para usuario de calidad"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
    Set Test_GetUserRole_UsuarioCalidad_DevuelveRolCalidad = testResult
End Function

' Prueba que GetUserRole devuelve Rol_Tecnico para usuario tÃ©cnico
Private Function Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockSolicitudRepository
    
    ' Nota: La nueva implementaciÃ³n de IConfig no permite SetValue
    ' Las pruebas deben usar la configuraciÃ³n real del sistema
    
    ' Crear recordset mock para usuario tÃ©cnico
    Dim mockRecordset As Object
    Set mockRecordset = CreateMockUserRecordset("No", "No", "SÃ­")
    mockRepository.SetMockRecordset mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize configService, mockLogger, mockRepository
    
    ' Act - Ejecutar el mÃ©todo bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("tecnico@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals Rol_Tecnico, userRole, "GetUserRole debe devolver Rol_Tecnico para usuario tÃ©cnico"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
    Set Test_GetUserRole_UsuarioTecnico_DevuelveRolTecnico = testResult
End Function

' Prueba que GetUserRole devuelve Rol_Desconocido para usuario no encontrado
Private Function Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockSolicitudRepository
    
    ' Nota: La nueva implementaciÃ³n de IConfig no permite SetValue
    ' Las pruebas deben usar la configuraciÃ³n real del sistema
    
    ' Crear recordset mock vacÃ­o (usuario no encontrado)
    Dim mockRecordset As Object
    Set mockRecordset = CreateEmptyUserRecordset()
    mockRepository.SetMockRecordset mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim authService As New CAuthService
    authService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Act - Ejecutar el mÃ©todo bajo prueba
    Dim userRole As E_UserRole
    userRole = authService.GetUserRole("inexistente@test.com")
    
    ' Assert - Verificar resultado
    modAssert.AssertEquals Rol_Desconocido, userRole, "GetUserRole debe devolver Rol_Desconocido para usuario no encontrado"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
    Set Test_GetUserRole_UsuarioDesconocido_DevuelveRolDesconocido = testResult
End Function

' ============================================================================
' FUNCIONES AUXILIARES PARA CREAR RECORDSETS MOCK
' ============================================================================

' Crea un recordset mock con datos de usuario para las pruebas de roles
Private Function CreateMockUserRecordset(esAdmin As String, esCalidad As String, esTecnico As String) As Object
    Dim rs As Object
    
    ' Crear recordset ADODB en memoria
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Definir campos del recordset que CAuthService espera de la Lanzadera
    rs.Fields.Append "EsAdministrador", 202, 10 ' adVarWChar
    rs.Fields.Append "EsUsuarioCalidad", 202, 10 ' adVarWChar
    rs.Fields.Append "EsUsuarioTecnico", 202, 10 ' adVarWChar
    
    ' Abrir recordset en memoria
    rs.Open
    
    ' AÃ±adir registro con datos del usuario
    rs.AddNew
    rs.Fields("EsAdministrador").Value = esAdmin
    rs.Fields("EsUsuarioCalidad").Value = esCalidad
    rs.Fields("EsUsuarioTecnico").Value = esTecnico
    rs.Update
    
    ' Mover al primer registro
    rs.MoveFirst
    
    Set CreateMockUserRecordset = rs
End Function

' Crea un recordset vacÃ­o para simular usuario no encontrado
Private Function CreateEmptyUserRecordset() As Object
    Dim rs As Object
    
    ' Crear recordset ADODB en memoria
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Definir campos bÃ¡sicos para el recordset vacÃ­o
    rs.Fields.Append "EsAdministrador", 202, 10 ' adVarWChar
    rs.Fields.Append "EsUsuarioCalidad", 202, 10 ' adVarWChar
    rs.Fields.Append "EsUsuarioTecnico", 202, 10 ' adVarWChar
    
    ' Abrir recordset en memoria (sin aÃ±adir registros, queda vacÃ­o)
    rs.Open
    
    Set CreateEmptyUserRecordset = rs
End Function

#End If











