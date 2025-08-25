Option Compare Database
Option Explicit
' ============================================================================
' MÃ³dulo: Test_DocumentService
' DescripciÃ³n: Suite de pruebas unitarias para DocumentService
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' VersiÃ³n: 2.0 - Estandarizado segÃºn framework
' ============================================================================

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function Test_DocumentService_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "DocumentService"
    
    ' Ejecutar todas las pruebas y aÃ±adir resultados
    suite.AddTestResult Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente()
    suite.AddTestResult Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia()
    suite.AddTestResult Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia()
    suite.AddTestResult Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud()
    suite.AddTestResult Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse()
    suite.AddTestResult Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente()
    
    Set Test_DocumentService_RunAll = suite
End Function

' ============================================================================
' PRUEBAS DE GENERACIÃ“N DE DOCUMENTOS
' ============================================================================

Private Function Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GenerarDocumento con datos vÃ¡lidos"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    
    ' Configurar mocks
    mockConfig.GetPlantillaPath_ReturnValue = "C:\Templates"
    mockConfig.GetGeneratedDocsPath_ReturnValue = "C:\Generated"
    mockConfig.IsTestMode_ReturnValue = True
    
    mockWordManager.AbrirDocumento_ReturnValue = True
    mockWordManager.ReemplazarTexto_ReturnValue = True
    mockWordManager.GuardarDocumento_ReturnValue = True
    
    ' Crear solicitud de prueba
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Permiso"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager
    
    ' Act
    Dim resultado As String
    resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    If resultado = "" Then
        testResult.Fail "No se generÃ³ el documento"
        GoTo Cleanup
    End If
    
    If Not mockWordManager.AbrirDocumento_WasCalled Then
        testResult.Fail "No se llamÃ³ a AbrirDocumento"
        GoTo Cleanup
    End If
    
    If Not mockWordManager.GuardarDocumento_WasCalled Then
        testResult.Fail "No se llamÃ³ a GuardarDocumento"
        GoTo Cleanup
    End If
    
    If Not mockWordManager.CerrarDocumento_WasCalled Then
        testResult.Fail "No se llamÃ³ a CerrarDocumento"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    Set Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GenerarDocumento con plantilla inexistente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    
    ' Configurar mocks para simular error
    mockConfig.GetPlantillaPath_ReturnValue = "C:\Templates"
    mockConfig.IsTestMode_ReturnValue = False ' Para que verifique existencia
    mockWordManager.AbrirDocumento_ReturnValue = False ' Simular fallo
    
    ' Crear solicitud de prueba
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Inexistente"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager
    
    ' Act
    Dim resultado As String
    resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    If resultado <> "" Then
        testResult.Fail "DeberÃ­a haber retornado cadena vacÃ­a"
    Else
        testResult.Pass
    End If
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    Set Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GenerarDocumento con error en WordManager"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    
    ' Configurar mocks
    mockConfig.GetPlantillaPath_ReturnValue = "C:\Templates"
    mockConfig.IsTestMode_ReturnValue = True
    mockWordManager.AbrirDocumento_ReturnValue = True
    mockWordManager.GuardarDocumento_ReturnValue = False ' Simular fallo al guardar
    
    ' Crear solicitud de prueba
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Permiso"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager
    
    ' Act
    Dim resultado As String
    resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    If resultado <> "" Then
        testResult.Fail "DeberÃ­a haber retornado cadena vacÃ­a"
    Else
        testResult.Pass
    End If
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    Set Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE LECTURA DE DOCUMENTOS
' ============================================================================

Private Function Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "LeerDocumento con documento vÃ¡lido"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    
    ' Configurar mocks
    Dim solicitudMock As New T_Solicitud
    solicitudMock.idSolicitud = 123
    solicitudMock.tipoSolicitud = "Permiso"
    
    mockRepository.GetSolicitudById_ReturnValue = solicitudMock
    mockRepository.UpdateSolicitud_ReturnValue = True
    mockWordManager.LeerContenidoDocumento_ReturnValue = "[nombre]Juan PÃ©rez[/nombre][fecha]2025-01-01[/fecha]"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager
    
    ' Act
    Dim resultado As Boolean
    resultado = docService.LeerDocumento("C:\test.docx", 123)
    
    ' Assert
    If Not resultado Then
        testResult.Fail "Fallo al leer documento"
        GoTo Cleanup
    End If
    
    If Not mockRepository.GetSolicitudById_WasCalled Then
        testResult.Fail "No se llamÃ³ a GetSolicitudById del repositorio"
        GoTo Cleanup
    End If
    
    If Not mockRepository.UpdateSolicitud_WasCalled Then
        testResult.Fail "No se llamÃ³ a UpdateSolicitud del repositorio"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    Set Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "LeerDocumento con documento inexistente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    
    ' Configurar mocks para simular error
    Dim solicitudMock As New T_Solicitud
    solicitudMock.idSolicitud = 123
    
    mockRepository.GetSolicitudById_ReturnValue = solicitudMock
    mockWordManager.LeerContenidoDocumento_ReturnValue = "" ' Simular fallo
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager
    
    ' Act
    Dim resultado As Boolean
    resultado = docService.LeerDocumento("C:\inexistente.docx", 123)
    
    ' Assert
    If resultado Then
        testResult.Fail "DeberÃ­a haber retornado False"
    Else
        testResult.Pass
    End If
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    Set Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE INICIALIZACIÃ“N
' ============================================================================

Private Function Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Initialize con dependencias vÃ¡lidas"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    
    ' Act
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager
    
    ' Assert
    ' Verificar que no se produzcan errores durante la inicializaciÃ³n
    testResult.Pass
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    Set Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado durante inicializaciÃ³n: " & Err.Description
    Resume Cleanup
End Function
