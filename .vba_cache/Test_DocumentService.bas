Attribute VB_Name = "Test_DocumentService"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_DocumentService
' Descripción: Suite de pruebas unitarias para DocumentService
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' Versión: 2.0 - Estandarizado según framework
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function Test_DocumentService_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "DocumentService"
    
    ' Ejecutar todas las pruebas y añadir resultados
    suite.AddTestResult Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente()
    suite.AddTestResult Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia()
    suite.AddTestResult Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia()
    suite.AddTestResult Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud()
    suite.AddTestResult Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse()
    suite.AddTestResult Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente()
    
    Set Test_DocumentService_RunAll = suite
End Function

' ============================================================================
' PRUEBAS DE GENERACIÓN DE DOCUMENTOS
' ============================================================================

Private Function Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GenerarDocumento con datos válidos"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    Dim mockMapeoRepository As New CMockMapeoRepository
    
    ' Configurar mocks
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "GENERATED_DOCS_PATH", "C:\Generated"
    mockConfig.AddSetting "IS_TEST_MODE", True
    
    mockWordManager.AddSetting "ABRIR_DOCUMENTO", True
    mockWordManager.AddSetting "REEMPLAZAR_TEXTO", True
    mockWordManager.AddSetting "GUARDAR_DOCUMENTO", True
    
    ' Crear solicitud de prueba
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Permiso"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository
    
    ' Act
    Dim Resultado As String
    Resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    If Resultado = "" Then
        testResult.Fail "No se generó el documento"
        GoTo Cleanup
    End If
    
    If Not mockWordManager.AbrirDocumento_WasCalled Then
        testResult.Fail "No se llamó a AbrirDocumento del WordManager"
        GoTo Cleanup
    End If
    
    If Not mockWordManager.GuardarDocumento_WasCalled Then
        testResult.Fail "No se llamó a GuardarDocumento del WordManager"
        GoTo Cleanup
    End If
    
    If Not mockWordManager.CerrarDocumento_WasCalled Then
        testResult.Fail "No se llamó a CerrarDocumento del WordManager"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    mockMapeoRepository.Reset
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
    Dim mockMapeoRepository As New CMockMapeoRepository
    
    ' Configurar mocks para simular error
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "IS_TEST_MODE", False ' Para que verifique existencia
    mockWordManager.AddSetting "ABRIR_DOCUMENTO", False ' Simular fallo
    
    ' Crear solicitud de prueba
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Inexistente"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository
    
    ' Act
    Dim Resultado As String
    Resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    If Resultado <> "" Then
        testResult.Fail "Debería haber retornado cadena vacía"
    Else
        testResult.Pass
    End If
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    mockMapeoRepository.Reset
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
    Dim mockMapeoRepository As New CMockMapeoRepository
    
    ' Configurar mocks
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "IS_TEST_MODE", True
    mockWordManager.AddSetting "ABRIR_DOCUMENTO", True
    mockWordManager.AddSetting "GUARDAR_DOCUMENTO", False ' Simular fallo al guardar
    
    ' Crear solicitud de prueba
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Permiso"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository
    
    ' Act
    Dim Resultado As String
    Resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    If Resultado <> "" Then
        testResult.Fail "Debería haber retornado cadena vacía"
    Else
        testResult.Pass
    End If
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    mockMapeoRepository.Reset
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
    testResult.Initialize "LeerDocumento con documento válido"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    Dim mockMapeoRepository As New CMockMapeoRepository
    
    ' Configurar mocks
    Dim solicitudMock As New T_Solicitud
    solicitudMock.idSolicitud = 123
    solicitudMock.tipoSolicitud = "Permiso"
    
    mockRepository.AddSetting "GET_SOLICITUD_BY_ID", solicitudMock
    mockRepository.AddSetting "SAVE_SOLICITUD", 1
    mockWordManager.AddSetting "LEER_CONTENIDO_DOCUMENTO", "[nombre]Juan Pérez[/nombre][fecha]2025-01-01[/fecha]"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository
    
    ' Act
    Dim Resultado As Boolean
    Resultado = docService.LeerDocumento("C:\test.docx", 123)
    
    ' Assert
    If Not Resultado Then
        testResult.Fail "Fallo al leer documento"
        GoTo Cleanup
    End If
    
    If Not mockRepository.GetSolicitudByIdCalled Then
        testResult.Fail "No se llamó a GetSolicitudById del repositorio"
        GoTo Cleanup
    End If
    
    If Not mockRepository.SaveSolicitudCalled Then
        testResult.Fail "No se llamó a SaveSolicitud del repositorio"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    mockMapeoRepository.Reset
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
    Dim mockMapeoRepository As New CMockMapeoRepository
    
    ' Configurar mocks para simular error
    Dim solicitudMock As New T_Solicitud
    solicitudMock.idSolicitud = 123
    
    mockRepository.AddSetting "GET_SOLICITUD_BY_ID", solicitudMock
    mockWordManager.AddSetting "LEER_CONTENIDO_DOCUMENTO", "" ' Simular fallo
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository
    
    ' Act
    Dim Resultado As Boolean
    Resultado = docService.LeerDocumento("C:\inexistente.docx", 123)
    
    ' Assert
    If Resultado Then
        testResult.Fail "Debería haber retornado False"
    Else
        testResult.Pass
    End If
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    mockMapeoRepository.Reset
    Set Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE INICIALIZACIÓN
' ============================================================================

Private Function Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Initialize con dependencias válidas"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim docService As New CDocumentService
    Dim mockConfig As New CMockConfig
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    Dim mockWordManager As New CMockWordManager
    Dim mockMapeoRepository As New CMockMapeoRepository
    
    ' Act
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository
    
    ' Assert
    ' Verificar que no se produzcan errores durante la inicialización
    testResult.Pass
    
Cleanup:
    ' Limpiar
    mockConfig.Reset
    mockRepository.Reset
    mockLogger.Reset
    mockWordManager.Reset
    mockMapeoRepository.Reset
    Set Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado durante inicialización: " & Err.Description
    Resume Cleanup
End Function


