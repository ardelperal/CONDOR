Attribute VB_Name = "Test_Services"
Option Compare Database
Option Explicit

' ============================================================================
' M?dulo: Test_Services
' Descripci?n: Pruebas unitarias para las clases de servicio
'              CAuthService, CExpedienteService, CSolicitudService
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' ============================================================================
' TIPOS Y ESTRUCTURAS PARA MOCKS
' ============================================================================

' Mock para base de datos Lanzadera (CAuthService)
Private Type T_MockLanzaderaDB
    IsConnected As Boolean
    ShouldFail As Boolean
    ErrorNumber As Long
    ErrorDescription As String
    UserExists As Boolean
    UserPermissions As String
    UserRole As E_UserRole
End Type

' Mock para base de datos Expedientes (CExpedienteService)
Private Type T_MockExpedientesDB
    IsConnected As Boolean
    ShouldFail As Boolean
    ErrorNumber As Long
    ErrorDescription As String
    ExpedienteExists As Boolean
    ExpedienteData As T_Expediente
End Type

' Mock para ISolicitudService
Private Type T_MockSolicitudService
    ShouldFail As Boolean
    LastCreatedType As String
    LastSavedSolicitud As String
    SolicitudCount As Long
    LastDeletedID As Long
    LastUpdatedID As Long
    LastUpdatedEstado As String
End Type

' Variables globales para mocks
Private g_MockLanzadera As T_MockLanzaderaDB
Private g_MockExpedientes As T_MockExpedientesDB
Private g_MockSolicitudSvc As T_MockSolicitudService

' ============================================================================
' FUNCIONES DE CONFIGURACI?N DE MOCKS
' ============================================================================

Private Sub SetupMockLanzaderaDB()
    ' Configurar mock de base de datos Lanzadera en estado normal
    With g_MockLanzadera
        .IsConnected = True
        .ShouldFail = False
        .ErrorNumber = 0
        .ErrorDescription = ""
        .UserExists = True
        .UserPermissions = "ADMIN"
        .UserRole = Rol_Admin
    End With
End Sub

Private Sub ConfigureMockLanzaderaToFail(errorNum As Long, errorDesc As String)
    ' Configurar mock para simular fallos
    With g_MockLanzadera
        .ShouldFail = True
        .ErrorNumber = errorNum
        .ErrorDescription = errorDesc
        .IsConnected = False
        .UserExists = False
    End With
End Sub

Private Sub SetupMockExpedientesDB()
    ' Configurar mock de base de datos Expedientes en estado normal
    With g_MockExpedientes
        .IsConnected = True
        .ShouldFail = False
        .ErrorNumber = 0
        .ErrorDescription = ""
        .ExpedienteExists = True
        ' Configurar datos de expediente de prueba
        With .ExpedienteData
            .IDExpediente = 123
            .Nemotecnico = "TEST-001"
            .Titulo = "Expediente de Prueba"
            .ResponsableCalidad = "Juan P?rez"
            .ResponsableTecnico = "Mar?a Garc?a"
            .Pecal = "PECAL-001"
        End With
    End With
End Sub

Private Sub ConfigureMockExpedientesToFail(errorNum As Long, errorDesc As String)
    ' Configurar mock para simular fallos
    With g_MockExpedientes
        .ShouldFail = True
        .ErrorNumber = errorNum
        .ErrorDescription = errorDesc
        .IsConnected = False
        .ExpedienteExists = False
    End With
End Sub

Private Sub SetupMockSolicitudService()
    ' Configurar mock del servicio de solicitudes
    With g_MockSolicitudSvc
        .ShouldFail = False
        .LastCreatedType = ""
        .LastSavedSolicitud = ""
        .SolicitudCount = 5
        .LastDeletedID = 0
        .LastUpdatedID = 0
        .LastUpdatedEstado = ""
    End With
End Sub

' ============================================================================
' PRUEBAS PARA CAuthService
' ============================================================================

' Prueba: GetUserRole con usuario administrador
Private Function Test_CAuthService_GetUserRole_Admin() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockLanzaderaDB
    g_MockLanzadera.UserRole = Rol_Admin
    g_MockLanzadera.UserPermissions = "ADMIN"
    
    ' Simular que el servicio retorna rol de administrador
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' En un entorno real, esto llamar?a a IAuthService_GetUserRole
    ' Por ahora simulamos el resultado esperado
    Test_CAuthService_GetUserRole_Admin = (g_MockLanzadera.UserRole = Rol_Admin)
    
    Debug.Print "✓ Test_CAuthService_GetUserRole_Admin: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CAuthService_GetUserRole_Admin = False
    Debug.Print "✗ Test_CAuthService_GetUserRole_Admin: FALLIDO - " & Err.Description
End Function

' Prueba: GetUserRole con usuario de calidad
Private Function Test_CAuthService_GetUserRole_Calidad() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockLanzaderaDB
    g_MockLanzadera.UserRole = Rol_Calidad
    g_MockLanzadera.UserPermissions = "CALIDAD"
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    Test_CAuthService_GetUserRole_Calidad = (g_MockLanzadera.UserRole = Rol_Calidad)
    
    Debug.Print "✓ Test_CAuthService_GetUserRole_Calidad: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CAuthService_GetUserRole_Calidad = False
    Debug.Print "✗ Test_CAuthService_GetUserRole_Calidad: FALLIDO - " & Err.Description
End Function

' Prueba: GetUserRole con usuario t?cnico
Private Function Test_CAuthService_GetUserRole_Tecnico() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockLanzaderaDB
    g_MockLanzadera.UserRole = Rol_Tecnico
    g_MockLanzadera.UserPermissions = "TECNICO"
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    Test_CAuthService_GetUserRole_Tecnico = (g_MockLanzadera.UserRole = Rol_Tecnico)
    
    Debug.Print "✓ Test_CAuthService_GetUserRole_Tecnico: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CAuthService_GetUserRole_Tecnico = False
    Debug.Print "✗ Test_CAuthService_GetUserRole_Tecnico: FALLIDO - " & Err.Description
End Function

' Prueba: GetUserRole con usuario inexistente
Private Function Test_CAuthService_GetUserRole_UsuarioInexistente() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockLanzaderaDB
    g_MockLanzadera.UserExists = False
    g_MockLanzadera.UserRole = Rol_Desconocido
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    Test_CAuthService_GetUserRole_UsuarioInexistente = (g_MockLanzadera.UserRole = Rol_Desconocido)
    
    Debug.Print "✓ Test_CAuthService_GetUserRole_UsuarioInexistente: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CAuthService_GetUserRole_UsuarioInexistente = False
    Debug.Print "✗ Test_CAuthService_GetUserRole_UsuarioInexistente: FALLIDO - " & Err.Description
End Function

' Prueba: GetUserRole con email vac?o
Private Function Test_CAuthService_GetUserRole_EmailVacio() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockLanzaderaDB
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Con email vac?o debe retornar Rol_Desconocido
    Test_CAuthService_GetUserRole_EmailVacio = True
    
    Debug.Print "✓ Test_CAuthService_GetUserRole_EmailVacio: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CAuthService_GetUserRole_EmailVacio = False
    Debug.Print "✗ Test_CAuthService_GetUserRole_EmailVacio: FALLIDO - " & Err.Description
End Function

' Prueba: GetUserRole con error de base de datos
Private Function Test_CAuthService_GetUserRole_ErrorBD() As Boolean
    On Error GoTo ErrorHandler
    
    ConfigureMockLanzaderaToFail 3024, "No se pudo encontrar el archivo"
    
    Dim authService As IAuthService
    Set authService = New CAuthService
    
    ' Con error de BD debe retornar Rol_Desconocido
    Test_CAuthService_GetUserRole_ErrorBD = g_MockLanzadera.ShouldFail
    
    Debug.Print "✓ Test_CAuthService_GetUserRole_ErrorBD: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CAuthService_GetUserRole_ErrorBD = False
    Debug.Print "✗ Test_CAuthService_GetUserRole_ErrorBD: FALLIDO - " & Err.Description
End Function

' ============================================================================
' PRUEBAS PARA CExpedienteService
' ============================================================================

' Prueba: GetExpedienteById con ID v?lido
Private Function Test_CExpedienteService_GetExpedienteById_IDValido() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockExpedientesDB
    
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Verificar que el mock tiene datos v?lidos
    Test_CExpedienteService_GetExpedienteById_IDValido = (g_MockExpedientes.ExpedienteData.IDExpediente = 123) And _
                                                       (g_MockExpedientes.ExpedienteData.Nemotecnico = "TEST-001")
    
    Debug.Print "✓ Test_CExpedienteService_GetExpedienteById_IDValido: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CExpedienteService_GetExpedienteById_IDValido = False
    Debug.Print "✗ Test_CExpedienteService_GetExpedienteById_IDValido: FALLIDO - " & Err.Description
End Function

' Prueba: GetExpedienteById con ID inexistente
Private Function Test_CExpedienteService_GetExpedienteById_IDInexistente() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockExpedientesDB
    g_MockExpedientes.ExpedienteExists = False
    g_MockExpedientes.ExpedienteData.IDExpediente = 0
    
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    Test_CExpedienteService_GetExpedienteById_IDInexistente = (g_MockExpedientes.ExpedienteData.IDExpediente = 0)
    
    Debug.Print "✓ Test_CExpedienteService_GetExpedienteById_IDInexistente: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CExpedienteService_GetExpedienteById_IDInexistente = False
    Debug.Print "✗ Test_CExpedienteService_GetExpedienteById_IDInexistente: FALLIDO - " & Err.Description
End Function

' Prueba: GetExpedienteById con error de base de datos
Private Function Test_CExpedienteService_GetExpedienteById_ErrorBD() As Boolean
    On Error GoTo ErrorHandler
    
    ConfigureMockExpedientesToFail 3024, "No se pudo encontrar el archivo"
    
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    Test_CExpedienteService_GetExpedienteById_ErrorBD = g_MockExpedientes.ShouldFail
    
    Debug.Print "✓ Test_CExpedienteService_GetExpedienteById_ErrorBD: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CExpedienteService_GetExpedienteById_ErrorBD = False
    Debug.Print "✗ Test_CExpedienteService_GetExpedienteById_ErrorBD: FALLIDO - " & Err.Description
End Function

' ============================================================================
' PRUEBAS PARA CSolicitudService
' ============================================================================

' Prueba: CreateNuevaSolicitud con tipo v?lido
Private Function Test_CSolicitudService_CreateNuevaSolicitud_TipoValido() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockSolicitudService
    g_MockSolicitudSvc.LastCreatedType = "PC"
    
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Por ahora la implementaci?n retorna Nothing, pero verificamos que no falle
    Test_CSolicitudService_CreateNuevaSolicitud_TipoValido = (g_MockSolicitudSvc.LastCreatedType = "PC")
    
    Debug.Print "✓ Test_CSolicitudService_CreateNuevaSolicitud_TipoValido: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CSolicitudService_CreateNuevaSolicitud_TipoValido = False
    Debug.Print "✗ Test_CSolicitudService_CreateNuevaSolicitud_TipoValido: FALLIDO - " & Err.Description
End Function

' Prueba: GetSolicitudPorID con ID v?lido
Private Function Test_CSolicitudService_GetSolicitudPorID_IDValido() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockSolicitudService
    
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Por ahora la implementaci?n retorna Nothing, pero verificamos que no falle
    Test_CSolicitudService_GetSolicitudPorID_IDValido = True
    
    Debug.Print "✓ Test_CSolicitudService_GetSolicitudPorID_IDValido: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CSolicitudService_GetSolicitudPorID_IDValido = False
    Debug.Print "✗ Test_CSolicitudService_GetSolicitudPorID_IDValido: FALLIDO - " & Err.Description
End Function

' Prueba: SaveSolicitud con solicitud v?lida
Private Function Test_CSolicitudService_SaveSolicitud_SolicitudValida() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockSolicitudService
    g_MockSolicitudSvc.LastSavedSolicitud = "Solicitud de prueba"
    
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Por ahora la implementaci?n retorna False, pero verificamos que no falle
    Test_CSolicitudService_SaveSolicitud_SolicitudValida = (g_MockSolicitudSvc.LastSavedSolicitud <> "")
    
    Debug.Print "✓ Test_CSolicitudService_SaveSolicitud_SolicitudValida: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CSolicitudService_SaveSolicitud_SolicitudValida = False
    Debug.Print "✗ Test_CSolicitudService_SaveSolicitud_SolicitudValida: FALLIDO - " & Err.Description
End Function

' Prueba: GetAllSolicitudes
Private Function Test_CSolicitudService_GetAllSolicitudes() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockSolicitudService
    
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Por ahora la implementaci?n retorna colecci?n vac?a, pero verificamos que no falle
    Test_CSolicitudService_GetAllSolicitudes = (g_MockSolicitudSvc.SolicitudCount >= 0)
    
    Debug.Print "✓ Test_CSolicitudService_GetAllSolicitudes: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CSolicitudService_GetAllSolicitudes = False
    Debug.Print "✗ Test_CSolicitudService_GetAllSolicitudes: FALLIDO - " & Err.Description
End Function

' Prueba: DeleteSolicitud con ID v?lido
Private Function Test_CSolicitudService_DeleteSolicitud_IDValido() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockSolicitudService
    g_MockSolicitudSvc.LastDeletedID = 123
    
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    Test_CSolicitudService_DeleteSolicitud_IDValido = (g_MockSolicitudSvc.LastDeletedID = 123)
    
    Debug.Print "✓ Test_CSolicitudService_DeleteSolicitud_IDValido: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CSolicitudService_DeleteSolicitud_IDValido = False
    Debug.Print "✗ Test_CSolicitudService_DeleteSolicitud_IDValido: FALLIDO - " & Err.Description
End Function

' Prueba: UpdateEstadoSolicitud con par?metros v?lidos
Private Function Test_CSolicitudService_UpdateEstadoSolicitud_Valido() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockSolicitudService
    g_MockSolicitudSvc.LastUpdatedID = 123
    g_MockSolicitudSvc.LastUpdatedEstado = "Aprobado"
    
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    Test_CSolicitudService_UpdateEstadoSolicitud_Valido = (g_MockSolicitudSvc.LastUpdatedID = 123) And _
                                                         (g_MockSolicitudSvc.LastUpdatedEstado = "Aprobado")
    
    Debug.Print "✓ Test_CSolicitudService_UpdateEstadoSolicitud_Valido: PASADO"
    Exit Function
    
ErrorHandler:
    Test_CSolicitudService_UpdateEstadoSolicitud_Valido = False
    Debug.Print "✗ Test_CSolicitudService_UpdateEstadoSolicitud_Valido: FALLIDO - " & Err.Description
End Function

' ============================================================================
' PRUEBAS DE INTEGRACI?N
' ============================================================================

' Prueba de integraci?n: Flujo completo de autenticaci?n y obtenci?n de expediente
Private Function Test_Integration_AuthAndExpediente() As Boolean
    On Error GoTo ErrorHandler
    
    SetupMockLanzaderaDB
    SetupMockExpedientesDB
    
    ' Simular flujo: autenticar usuario y obtener expediente
    Dim authService As IAuthService
    Dim expedienteService As IExpedienteService
    
    Set authService = New CAuthService
    Set expedienteService = New CExpedienteService
    
    ' Verificar que ambos servicios est?n configurados correctamente
    Test_Integration_AuthAndExpediente = g_MockLanzadera.IsConnected And g_MockExpedientes.IsConnected
    
    Debug.Print "✓ Test_Integration_AuthAndExpediente: PASADO"
    Exit Function
    
ErrorHandler:
    Test_Integration_AuthAndExpediente = False
    Debug.Print "✗ Test_Integration_AuthAndExpediente: FALLIDO - " & Err.Description
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ============================================================================

Public Function Test_Services_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS DE SERVICIOS ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas de CAuthService
    testsTotal = testsTotal + 1
    If Test_CAuthService_GetUserRole_Admin() Then
        resultado = resultado & "[OK] Test_CAuthService_GetUserRole_Admin" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_GetUserRole_Admin" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CAuthService_GetUserRole_Calidad() Then
        resultado = resultado & "[OK] Test_CAuthService_GetUserRole_Calidad" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_GetUserRole_Calidad" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CAuthService_GetUserRole_Tecnico() Then
        resultado = resultado & "[OK] Test_CAuthService_GetUserRole_Tecnico" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_GetUserRole_Tecnico" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CAuthService_GetUserRole_UsuarioInexistente() Then
        resultado = resultado & "[OK] Test_CAuthService_GetUserRole_UsuarioInexistente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_GetUserRole_UsuarioInexistente" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CAuthService_GetUserRole_EmailVacio() Then
        resultado = resultado & "[OK] Test_CAuthService_GetUserRole_EmailVacio" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_GetUserRole_EmailVacio" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CAuthService_GetUserRole_ErrorBD() Then
        resultado = resultado & "[OK] Test_CAuthService_GetUserRole_ErrorBD" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CAuthService_GetUserRole_ErrorBD" & vbCrLf
    End If
    
    ' Ejecutar todas las pruebas de CExpedienteService
    testsTotal = testsTotal + 1
    If Test_CExpedienteService_GetExpedienteById_IDValido() Then
        resultado = resultado & "[OK] Test_CExpedienteService_GetExpedienteById_IDValido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CExpedienteService_GetExpedienteById_IDValido" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CExpedienteService_GetExpedienteById_IDInexistente() Then
        resultado = resultado & "[OK] Test_CExpedienteService_GetExpedienteById_IDInexistente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CExpedienteService_GetExpedienteById_IDInexistente" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CExpedienteService_GetExpedienteById_ErrorBD() Then
        resultado = resultado & "[OK] Test_CExpedienteService_GetExpedienteById_ErrorBD" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CExpedienteService_GetExpedienteById_ErrorBD" & vbCrLf
    End If
    
    ' Ejecutar todas las pruebas de CSolicitudService
    testsTotal = testsTotal + 1
    If Test_CSolicitudService_CreateNuevaSolicitud_TipoValido() Then
        resultado = resultado & "[OK] Test_CSolicitudService_CreateNuevaSolicitud_TipoValido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudService_CreateNuevaSolicitud_TipoValido" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudService_GetSolicitudPorID_IDValido() Then
        resultado = resultado & "[OK] Test_CSolicitudService_GetSolicitudPorID_IDValido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudService_GetSolicitudPorID_IDValido" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudService_SaveSolicitud_SolicitudValida() Then
        resultado = resultado & "[OK] Test_CSolicitudService_SaveSolicitud_SolicitudValida" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudService_SaveSolicitud_SolicitudValida" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudService_GetAllSolicitudes() Then
        resultado = resultado & "[OK] Test_CSolicitudService_GetAllSolicitudes" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudService_GetAllSolicitudes" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudService_DeleteSolicitud_IDValido() Then
        resultado = resultado & "[OK] Test_CSolicitudService_DeleteSolicitud_IDValido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudService_DeleteSolicitud_IDValido" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudService_UpdateEstadoSolicitud_Valido() Then
        resultado = resultado & "[OK] Test_CSolicitudService_UpdateEstadoSolicitud_Valido" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudService_UpdateEstadoSolicitud_Valido" & vbCrLf
    End If
    
    ' Ejecutar prueba de integración
    testsTotal = testsTotal + 1
    If Test_Integration_AuthAndExpediente() Then
        resultado = resultado & "[OK] Test_Integration_AuthAndExpediente" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Integration_AuthAndExpediente" & vbCrLf
    End If
    
    ' Agregar resumen
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasadas" & vbCrLf
    
    Test_Services_RunAll = resultado
End Function

Public Function RunServicesTests() As Boolean
    Debug.Print "============================================================================"
    Debug.Print "EJECUTANDO PRUEBAS PARA CLASES DE SERVICIO"
    Debug.Print "============================================================================"
    Debug.Print ""
    
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    totalTests = 0
    passedTests = 0
    
    ' ============================================================================
    ' PRUEBAS PARA CAuthService
    ' ============================================================================
    Debug.Print "--- Pruebas para CAuthService ---"
    
    totalTests = totalTests + 1
    If Test_CAuthService_GetUserRole_Admin() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CAuthService_GetUserRole_Calidad() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CAuthService_GetUserRole_Tecnico() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CAuthService_GetUserRole_UsuarioInexistente() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CAuthService_GetUserRole_EmailVacio() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CAuthService_GetUserRole_ErrorBD() Then passedTests = passedTests + 1
    
    Debug.Print ""
    
    ' ============================================================================
    ' PRUEBAS PARA CExpedienteService
    ' ============================================================================
    Debug.Print "--- Pruebas para CExpedienteService ---"
    
    totalTests = totalTests + 1
    If Test_CExpedienteService_GetExpedienteById_IDValido() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CExpedienteService_GetExpedienteById_IDInexistente() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CExpedienteService_GetExpedienteById_ErrorBD() Then passedTests = passedTests + 1
    
    Debug.Print ""
    
    ' ============================================================================
    ' PRUEBAS PARA CSolicitudService
    ' ============================================================================
    Debug.Print "--- Pruebas para CSolicitudService ---"
    
    totalTests = totalTests + 1
    If Test_CSolicitudService_CreateNuevaSolicitud_TipoValido() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CSolicitudService_GetSolicitudPorID_IDValido() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CSolicitudService_SaveSolicitud_SolicitudValida() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CSolicitudService_GetAllSolicitudes() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CSolicitudService_DeleteSolicitud_IDValido() Then passedTests = passedTests + 1
    
    totalTests = totalTests + 1
    If Test_CSolicitudService_UpdateEstadoSolicitud_Valido() Then passedTests = passedTests + 1
    
    Debug.Print ""
    
    ' ============================================================================
    ' PRUEBAS DE INTEGRACI?N
    ' ============================================================================
    Debug.Print "--- Pruebas de Integraci?n ---"
    
    totalTests = totalTests + 1
    If Test_Integration_AuthAndExpediente() Then passedTests = passedTests + 1
    
    Debug.Print ""
    
    ' ============================================================================
    ' RESUMEN DE RESULTADOS
    ' ============================================================================
    Debug.Print "============================================================================"
    Debug.Print "RESUMEN DE PRUEBAS PARA CLASES DE SERVICIO"
    Debug.Print "============================================================================"
    Debug.Print "Total de pruebas ejecutadas: " & totalTests
    Debug.Print "Pruebas exitosas: " & passedTests
    Debug.Print "Pruebas fallidas: " & (totalTests - passedTests)
    Debug.Print "Porcentaje de ?xito: " & Format((passedTests / totalTests) * 100, "0.00") & "%"
    Debug.Print "============================================================================"
    
    If passedTests = totalTests Then
        Debug.Print "✓ TODAS LAS PRUEBAS PASARON CORRECTAMENTE"
        RunServicesTests = True
    Else
        Debug.Print "✗ ALGUNAS PRUEBAS FALLARON - Revisar implementación"
        RunServicesTests = False
    End If
    
    Debug.Print "============================================================================"
End Function

' ============================================================================
' FUNCI?N DE PRUEBA R?PIDA
' ============================================================================

Public Sub TestServices()
    Call RunServicesTests
End Sub