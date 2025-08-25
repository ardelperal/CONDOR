Attribute VB_Name = "Test_CSolicitudRepository"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' SUITE DE PRUEBAS DE INTEGRACIÃ“N PARA CSolicitudRepository
' Arquitectura: Pruebas Reales con ConexiÃ³n al Backend
' Version: 2.0 - ReconstrucciÃ³n Total
' ============================================================================
' Pruebas de integraciÃ³n que validan las operaciones de CSolicitudRepository
' contra la base de datos real del backend.
' ============================================================================

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function RunAllTests() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CSolicitudRepository - Pruebas de IntegraciÃ³n"
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult Test_SaveSolicitud_Y_GetSolicitudById_CicloCompleto()
    suiteResult.AddTestResult Test_GetSolicitudById_ConIdInexistente_DebeRetornarNothing()
    suiteResult.AddTestResult Test_SaveSolicitud_ConSolicitudNueva_DebeAsignarId()
    suiteResult.AddTestResult Test_SaveSolicitud_ConSolicitudExistente_DebeActualizar()
    suiteResult.AddTestResult Test_Repository_SinInicializar_DebeLanzarError()
    
    Set RunAllTests = suiteResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÃ“N
' ============================================================================

Private Function Test_SaveSolicitud_Y_GetSolicitudById_CicloCompleto() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud y GetSolicitudById - Ciclo completo de guardado y recuperaciÃ³n"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim repository As New CSolicitudRepository
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "desarrollo2024"
    
    Dim mockLogger As New CMockOperationLogger
    
    ' Inicializar dependencias
    repository.Initialize mockConfig, mockLogger
    
    ' Crear solicitud de prueba
    Dim solicitudOriginal As New T_Solicitud
    With solicitudOriginal
        .idSolicitud = 0 ' Nueva solicitud
        .idExpediente = "TEST-EXP-" & Format(Now(), "yyyymmddhhnnss")
        .tipoSolicitud = "PC"
        .idEstadoInterno = 1 ' ID del estado Borrador
        .fechaCreacion = Now()
        .usuarioCreacion = "TEST_USER"
        .observaciones = "Solicitud de prueba de integraciÃ³n"
    End With
    
    Dim idGenerado As Long
    Dim solicitudRecuperada As T_Solicitud
    
    ' Act - Guardar solicitud
    idGenerado = repository.SaveSolicitud(solicitudOriginal)
    
    ' Assert - Verificar que se asignÃ³ un ID
    If idGenerado <= 0 Then
        testResult.Fail "SaveSolicitud debe retornar un ID vÃ¡lido, pero retornÃ³ " & idGenerado
        GoTo Cleanup
    End If
    
    ' Act - Recuperar solicitud
    Set solicitudRecuperada = repository.GetSolicitudById(idGenerado)
    
    ' Assert - Verificar que se recuperÃ³ correctamente
    If solicitudRecuperada Is Nothing Then
        testResult.Fail "GetSolicitudById debe retornar la solicitud guardada"
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.idSolicitud <> idGenerado Then
        testResult.Fail "El ID de la solicitud recuperada debe ser " & idGenerado & ", pero fue " & solicitudRecuperada.idSolicitud
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.idExpediente <> solicitudOriginal.idExpediente Then
        testResult.Fail "El ID del expediente debe coincidir: esperado '" & solicitudOriginal.idExpediente & "', obtenido '" & solicitudRecuperada.idExpediente & "'"
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.tipoSolicitud <> solicitudOriginal.tipoSolicitud Then
        testResult.Fail "El tipo de solicitud debe coincidir: esperado '" & solicitudOriginal.tipoSolicitud & "', obtenido '" & solicitudRecuperada.tipoSolicitud & "'"
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.idEstadoInterno <> solicitudOriginal.idEstadoInterno Then
        testResult.Fail "El estado interno debe coincidir: esperado '" & solicitudOriginal.idEstadoInterno & "', obtenido '" & solicitudRecuperada.idEstadoInterno & "'"
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.usuarioCreacion <> solicitudOriginal.usuarioCreacion Then
        testResult.Fail "El usuario de creaciÃ³n debe coincidir: esperado '" & solicitudOriginal.usuarioCreacion & "', obtenido '" & solicitudRecuperada.usuarioCreacion & "'"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    ' Limpiar datos de prueba
    If idGenerado > 0 Then
        Call LimpiarSolicitudDePrueba(idGenerado, mockConfig)
    End If
    
    Set Test_SaveSolicitud_Y_GetSolicitudById_CicloCompleto = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_GetSolicitudById_ConIdInexistente_DebeRetornarNothing() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetSolicitudById con ID inexistente debe retornar Nothing"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim repository As New CSolicitudRepository
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "desarrollo2024"
    
    Dim mockLogger As New CMockOperationLogger
    
    repository.Initialize mockConfig, mockLogger
    
    ' Act - Buscar solicitud con ID que no existe
    Dim idInexistente As Long
    idInexistente = -999999 ' ID que seguramente no existe
    
    Dim resultado As T_Solicitud
    Set resultado = repository.GetSolicitudById(idInexistente)
    
    ' Assert
    If Not (resultado Is Nothing) Then
        testResult.Fail "GetSolicitudById con ID inexistente debe retornar Nothing"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set Test_GetSolicitudById_ConIdInexistente_DebeRetornarNothing = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_SaveSolicitud_ConSolicitudNueva_DebeAsignarId() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud con solicitud nueva debe asignar ID automÃ¡ticamente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim repository As New CSolicitudRepository
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "desarrollo2024"
    
    Dim mockLogger As New CMockOperationLogger
    
    repository.Initialize mockConfig, mockLogger
    
    Dim solicitud As New T_Solicitud
    With solicitud
        .idSolicitud = 0 ' Nueva solicitud
        .idExpediente = "TEST-NEW-" & Format(Now(), "yyyymmddhhnnss")
        .tipoSolicitud = "PC"
        .idEstadoInterno = 1 ' ID del estado Borrador
        .fechaCreacion = Now()
        .usuarioCreacion = "TEST_USER"
    End With
    
    ' Act
    Dim idGenerado As Long
    idGenerado = repository.SaveSolicitud(solicitud)
    
    ' Assert
    If idGenerado <= 0 Then
        testResult.Fail "SaveSolicitud debe asignar un ID positivo, pero asignÃ³ " & idGenerado
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    ' Limpiar datos de prueba
    If idGenerado > 0 Then
        Call LimpiarSolicitudDePrueba(idGenerado, mockConfig)
    End If
    
    Set Test_SaveSolicitud_ConSolicitudNueva_DebeAsignarId = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_SaveSolicitud_ConSolicitudExistente_DebeActualizar() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud con solicitud existente debe actualizar los datos"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim repository As New CSolicitudRepository
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "desarrollo2024"
    
    Dim mockLogger As New CMockOperationLogger
    
    repository.Initialize mockConfig, mockLogger
    
    ' Crear y guardar solicitud inicial
    Dim solicitudInicial As New T_Solicitud
    With solicitudInicial
        .idSolicitud = 0
        .idExpediente = "TEST-UPD-" & Format(Now(), "yyyymmddhhnnss")
        .tipoSolicitud = "PC"
        .idEstadoInterno = 1 ' ID del estado Borrador
        .fechaCreacion = Now()
        .usuarioCreacion = "TEST_USER"
        .observaciones = "ObservaciÃ³n inicial"
    End With
    
    Dim idGenerado As Long
    idGenerado = repository.SaveSolicitud(solicitudInicial)
    
    If idGenerado <= 0 Then
        testResult.Fail "No se pudo crear la solicitud inicial para la prueba"
        GoTo Cleanup
    End If
    
    ' Modificar la solicitud
    Dim solicitudModificada As New T_Solicitud
    With solicitudModificada
        .idSolicitud = idGenerado
        .idExpediente = solicitudInicial.idExpediente
        .tipoSolicitud = "PC"
        .idEstadoInterno = 2 ' ID del estado En Proceso
        .fechaCreacion = solicitudInicial.fechaCreacion
        .usuarioCreacion = solicitudInicial.usuarioCreacion
        .fechaModificacion = Now()
        .usuarioModificacion = "TEST_USER_MOD"
        .observaciones = "ObservaciÃ³n modificada" ' Cambiar observaciones
    End With
    
    ' Act - Actualizar solicitud
    Dim idActualizado As Long
    idActualizado = repository.SaveSolicitud(solicitudModificada)
    
    ' Assert - Verificar que retorna el mismo ID
    If idActualizado <> idGenerado Then
        testResult.Fail "SaveSolicitud debe retornar el mismo ID para actualizaciones: esperado " & idGenerado & ", obtenido " & idActualizado
        GoTo Cleanup
    End If
    
    ' Recuperar y verificar cambios
    Dim solicitudRecuperada As T_Solicitud
    Set solicitudRecuperada = repository.GetSolicitudById(idGenerado)
    
    If solicitudRecuperada Is Nothing Then
        testResult.Fail "No se pudo recuperar la solicitud actualizada"
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.idEstadoInterno <> 2 Then
        testResult.Fail "El estado interno debe haberse actualizado a ID 2 (En Proceso), pero es '" & solicitudRecuperada.idEstadoInterno & "'"
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.observaciones <> "ObservaciÃ³n modificada" Then
        testResult.Fail "Las observaciones deben haberse actualizado a 'ObservaciÃ³n modificada', pero son '" & solicitudRecuperada.observaciones & "'"
        GoTo Cleanup
    End If
    
    If solicitudRecuperada.usuarioModificacion <> "TEST_USER_MOD" Then
        testResult.Fail "El usuario de modificaciÃ³n debe ser 'TEST_USER_MOD', pero es '" & solicitudRecuperada.usuarioModificacion & "'"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    ' Limpiar datos de prueba
    If idGenerado > 0 Then
        Call LimpiarSolicitudDePrueba(idGenerado, mockConfig)
    End If
    
    Set Test_SaveSolicitud_ConSolicitudExistente_DebeActualizar = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_Repository_SinInicializar_DebeLanzarError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Repository sin inicializar debe lanzar error"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim repository As New CSolicitudRepository
    ' No inicializar el repositorio
    
    Dim solicitud As New T_Solicitud
    solicitud.idSolicitud = 0
    
    ' Act & Assert - Intentar usar SaveSolicitud sin inicializar
    Dim errorOcurred As Boolean
    errorOcurred = False
    
    On Error Resume Next
    Dim resultado As Long
    resultado = repository.SaveSolicitud(solicitud)
    If Err.Number <> 0 Then errorOcurred = True
    On Error GoTo ErrorHandler
    
    If Not errorOcurred Then
        testResult.Fail "SaveSolicitud debe lanzar un error cuando el repositorio no estÃ¡ inicializado"
        GoTo Cleanup
    End If
    
    ' Act & Assert - Intentar usar GetSolicitudById sin inicializar
    errorOcurred = False
    
    On Error Resume Next
    Dim solicitudResult As T_Solicitud
    Set solicitudResult = repository.GetSolicitudById(1)
    If Err.Number <> 0 Then errorOcurred = True
    On Error GoTo ErrorHandler
    
    If Not errorOcurred Then
        testResult.Fail "GetSolicitudById debe lanzar un error cuando el repositorio no estÃ¡ inicializado"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set Test_Repository_SinInicializar_DebeLanzarError = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' MÃ‰TODOS AUXILIARES PARA LIMPIEZA
' ============================================================================

Private Sub LimpiarSolicitudDePrueba(ByVal idSolicitud As Long, ByRef mockConfig As CMockConfig)
    On Error Resume Next
    
    ' Conectar al backend para eliminar la solicitud de prueba
    Dim db As DAO.Database
    Dim rutaBackend As String
    Dim passwordBackend As String
    
    rutaBackend = mockConfig.GetDataPath()
    passwordBackend = mockConfig.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(rutaBackend, False, False, ";PWD=" & passwordBackend)
    
    ' Eliminar la solicitud de prueba
    Dim qdf As DAO.QueryDef
    Set qdf = db.CreateQueryDef("", "DELETE FROM T_Solicitud WHERE idSolicitud = ?")
    qdf.Parameters(0) = idSolicitud
    qdf.Execute
    
    ' Limpiar recursos
    qdf.Close
    db.Close
    Set qdf = Nothing
    Set db = Nothing
    
    On Error GoTo 0
End Sub

#End If
