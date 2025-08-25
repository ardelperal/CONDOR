Attribute VB_Name = "Test_Solicitud"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CSolicitudService
' Arquitectura: Pruebas Aisladas con Inyección de Dependencias y Mocks
' Version: 2.0 - Reconstrucción Total
' ============================================================================
' Pruebas unitarias que validan la lógica de negocio de CSolicitudService
' usando mocks para aislar las dependencias externas.
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function RunAllTests() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_Solicitud - Pruebas Unitarias CSolicitudService"
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult Test_CreateSolicitud_ConParametrosValidos_DebeCrearSolicitudConEstadoBorrador()
    suiteResult.AddTestResult Test_CreateSolicitud_ConIdExpedienteVacio_DebeLanzarError()
    suiteResult.AddTestResult Test_CreateSolicitud_ConTipoVacio_DebeLanzarError()
    suiteResult.AddTestResult Test_CreateSolicitud_SinInicializar_DebeLanzarError()
    suiteResult.AddTestResult Test_SaveSolicitud_ConSolicitudValida_DebeActualizarFechaModificacion()
    suiteResult.AddTestResult Test_SaveSolicitud_ConSolicitudNula_DebeLanzarError()
    suiteResult.AddTestResult Test_SaveSolicitud_SinInicializar_DebeLanzarError()
    
    Set RunAllTests = suiteResult
End Function

' ============================================================================
' PRUEBAS DE CreateSolicitud
' ============================================================================

Private Function Test_CreateSolicitud_ConParametrosValidos_DebeCrearSolicitudConEstadoBorrador() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateSolicitud con parámetros válidos debe crear solicitud con estado Borrador"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim solicitudService As New CSolicitudService
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    
    ' Configurar mocks
    mockRepository.SetSaveSolicitudReturnValue 123 ' ID de la nueva solicitud
    
    ' Inicializar servicio con dependencias
    solicitudService.Initialize mockRepository, mockLogger
    
    ' Act
    Dim resultado As T_Solicitud
    Set resultado = solicitudService.CreateSolicitud("EXP-2024-001", "PC")
    
    ' Assert
    If resultado Is Nothing Then
        testResult.Fail "El resultado no debe ser Nothing"
        GoTo Cleanup
    End If
    
    If resultado.idSolicitud <> 123 Then
        testResult.Fail "El ID de la solicitud debe ser 123, pero fue " & resultado.idSolicitud
        GoTo Cleanup
    End If
    
    If resultado.idExpediente <> "EXP-2024-001" Then
        testResult.Fail "El ID del expediente debe ser 'EXP-2024-001', pero fue '" & resultado.idExpediente & "'"
        GoTo Cleanup
    End If
    
    If resultado.tipoSolicitud <> "PC" Then
        testResult.Fail "El tipo de solicitud debe ser 'PC', pero fue '" & resultado.tipoSolicitud & "'"
        GoTo Cleanup
    End If
    
    If resultado.estadoInterno <> "Borrador" Then
        testResult.Fail "El estado interno debe ser 'Borrador', pero fue '" & resultado.estadoInterno & "'"
        GoTo Cleanup
    End If
    
    If Not mockRepository.SaveSolicitudCalled Then
        testResult.Fail "Debe llamarse a SaveSolicitud del repositorio"
        GoTo Cleanup
    End If
    
    If Not mockLogger.LogOperationCalled Then
        testResult.Fail "Debe llamarse a LogOperation del logger"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    mockRepository.Reset
    mockLogger.ClearLog
    Set Test_CreateSolicitud_ConParametrosValidos_DebeCrearSolicitudConEstadoBorrador = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_CreateSolicitud_ConIdExpedienteVacio_DebeLanzarError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateSolicitud con ID expediente vacío debe lanzar error"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim solicitudService As New CSolicitudService
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    
    solicitudService.Initialize mockRepository, mockLogger
    
    ' Act & Assert
    Dim errorOcurred As Boolean
    errorOcurred = False
    
    On Error Resume Next
    Dim resultado As T_Solicitud
    Set resultado = solicitudService.CreateSolicitud("", "PC")
    If Err.Number <> 0 Then errorOcurred = True
    On Error GoTo ErrorHandler
    
    If Not errorOcurred Then
        testResult.Fail "Debe lanzar un error cuando idExpediente está vacío"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    mockRepository.Reset
    mockLogger.ClearLog
    Set Test_CreateSolicitud_ConIdExpedienteVacio_DebeLanzarError = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_CreateSolicitud_ConTipoVacio_DebeLanzarError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateSolicitud con tipo vacío debe lanzar error"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim solicitudService As New CSolicitudService
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    
    solicitudService.Initialize mockRepository, mockLogger
    
    ' Act & Assert
    Dim errorOcurred As Boolean
    errorOcurred = False
    
    On Error Resume Next
    Dim resultado As T_Solicitud
    Set resultado = solicitudService.CreateSolicitud("EXP-2024-001", "")
    If Err.Number <> 0 Then errorOcurred = True
    On Error GoTo ErrorHandler
    
    If Not errorOcurred Then
        testResult.Fail "Debe lanzar un error cuando tipo está vacío"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    mockRepository.Reset
    mockLogger.ClearLog
    Set Test_CreateSolicitud_ConTipoVacio_DebeLanzarError = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_CreateSolicitud_SinInicializar_DebeLanzarError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateSolicitud sin inicializar debe lanzar error"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim solicitudService As New CSolicitudService
    ' No inicializar el servicio
    
    ' Act & Assert
    Dim errorOcurred As Boolean
    errorOcurred = False
    
    On Error Resume Next
    Dim resultado As T_Solicitud
    Set resultado = solicitudService.CreateSolicitud("EXP-2024-001", "PC")
    If Err.Number <> 0 Then errorOcurred = True
    On Error GoTo ErrorHandler
    
    If Not errorOcurred Then
        testResult.Fail "Debe lanzar un error cuando el servicio no está inicializado"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set Test_CreateSolicitud_SinInicializar_DebeLanzarError = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE SaveSolicitud
' ============================================================================

Private Function Test_SaveSolicitud_ConSolicitudValida_DebeActualizarFechaModificacion() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud con solicitud válida debe actualizar fecha de modificación"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim solicitudService As New CSolicitudService
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    
    ' Configurar mocks
    mockRepository.SetSaveSolicitudReturnValue 456 ' ID exitoso
    
    solicitudService.Initialize mockRepository, mockLogger
    
    ' Crear solicitud de prueba
    Dim solicitud As New T_Solicitud
    With solicitud
        .idSolicitud = 456
        .idExpediente = "EXP-2024-002"
        .tipoSolicitud = "PC"
        .estadoInterno = "Borrador"
        .fechaCreacion = DateAdd("d", -1, Now()) ' Ayer
        .fechaModificacion = Null ' Sin modificación previa
    End With
    
    ' Act
    Dim resultado As Boolean
    resultado = solicitudService.SaveSolicitud(solicitud)
    
    ' Assert
    If Not resultado Then
        testResult.Fail "SaveSolicitud debe retornar True para una solicitud válida"
        GoTo Cleanup
    End If
    
    If Not mockRepository.SaveSolicitudCalled Then
        testResult.Fail "Debe llamarse a SaveSolicitud del repositorio"
        GoTo Cleanup
    End If
    
    If IsNull(solicitud.fechaModificacion) Then
        testResult.Fail "La fecha de modificación debe ser actualizada"
        GoTo Cleanup
    End If
    
    If Len(Trim(solicitud.usuarioModificacion)) = 0 Then
        testResult.Fail "El usuario de modificación debe ser establecido"
        GoTo Cleanup
    End If
    
    If Not mockLogger.LogOperationCalled Then
        testResult.Fail "Debe llamarse a LogOperation del logger"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    mockRepository.Reset
    mockLogger.ClearLog
    Set Test_SaveSolicitud_ConSolicitudValida_DebeActualizarFechaModificacion = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_SaveSolicitud_ConSolicitudNula_DebeLanzarError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud con solicitud nula debe lanzar error"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim solicitudService As New CSolicitudService
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    
    solicitudService.Initialize mockRepository, mockLogger
    
    ' Act & Assert
    Dim errorOcurred As Boolean
    errorOcurred = False
    
    On Error Resume Next
    Dim resultado As Boolean
    resultado = solicitudService.SaveSolicitud(Nothing)
    If Err.Number <> 0 Then errorOcurred = True
    On Error GoTo ErrorHandler
    
    If Not errorOcurred Then
        testResult.Fail "Debe lanzar un error cuando la solicitud es Nothing"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    mockRepository.Reset
    mockLogger.ClearLog
    Set Test_SaveSolicitud_ConSolicitudNula_DebeLanzarError = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_SaveSolicitud_SinInicializar_DebeLanzarError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud sin inicializar debe lanzar error"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim solicitudService As New CSolicitudService
    ' No inicializar el servicio
    
    Dim solicitud As New T_Solicitud
    solicitud.idSolicitud = 123
    
    ' Act & Assert
    Dim errorOcurred As Boolean
    errorOcurred = False
    
    On Error Resume Next
    Dim resultado As Boolean
    resultado = solicitudService.SaveSolicitud(solicitud)
    If Err.Number <> 0 Then errorOcurred = True
    On Error GoTo ErrorHandler
    
    If Not errorOcurred Then
        testResult.Fail "Debe lanzar un error cuando el servicio no está inicializado"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set Test_SaveSolicitud_SinInicializar_DebeLanzarError = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

#End If