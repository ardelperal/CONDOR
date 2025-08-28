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

' Variables a nivel de módulo para los mocks
Private docService As CDocumentService
Private mockConfig As CMockConfig
Private mockRepository As CMockSolicitudRepository
Private mockLogger As CMockOperationLogger
Private mockWordManager As CMockWordManager
Private mockMapeoRepository As CMockMapeoRepository
Private mockErrorHandler As CMockErrorHandlerService

' =====================================================
' SETUP Y TEARDOWN
' =====================================================

Private Sub Setup()
    ' Instanciar todos los mocks necesarios para las pruebas
    Set docService = New CDocumentService
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
End Sub

Private Sub Teardown()
    ' Limpiar y resetear todos los mocks para aislar las pruebas
    If Not mockConfig Is Nothing Then mockConfig.Reset
    If Not mockRepository Is Nothing Then mockRepository.Reset
    If Not mockLogger Is Nothing Then mockLogger.Reset
    If Not mockWordManager Is Nothing Then mockWordManager.Reset
    If Not mockMapeoRepository Is Nothing Then mockMapeoRepository.Reset
    If Not mockErrorHandler Is Nothing Then mockErrorHandler.Reset

    Set docService = Nothing
    Set mockConfig = Nothing
    Set mockRepository = Nothing
    Set mockLogger = Nothing
    Set mockWordManager = Nothing
    Set mockMapeoRepository = Nothing
    Set mockErrorHandler = Nothing
End Sub

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function Test_DocumentService_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    Call suite.Initialize("DocumentService")
    
    ' Ejecutar todas las pruebas y añadir resultados
    Call suite.AddTestResult(Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente())
    Call suite.AddTestResult(Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia())
    Call suite.AddTestResult(Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia())
    Call suite.AddTestResult(Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud())
    Call suite.AddTestResult(Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse())
    Call suite.AddTestResult(Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente())
    
    Set Test_DocumentService_RunAll = suite
End Function

' ============================================================================
' PRUEBAS DE GENERACIÓN DE DOCUMENTOS
' ============================================================================

Private Function Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente() As CTestResult
    Set Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente = New CTestResult
    Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente.Initialize "GenerarDocumento con datos válidos"
    On Error GoTo TestFail

    ' Arrange
    Call Setup ' Asegura que todos los mocks están instanciados

    ' Configurar mocks para un escenario de éxito
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "GENERATED_DOCS_PATH", "C:\Generated"
    mockConfig.AddSetting "IS_TEST_MODE", True

    mockWordManager.ConfigureAbrirDocumento True
    mockWordManager.ConfigureReemplazarTexto True
    mockWordManager.ConfigureGuardarDocumento True

    Dim solicitudPrueba As New ESolicitud
    solicitudPrueba.idSolicitud = 123
    solicitudPrueba.codigoSolicitud = "SOL001"
    solicitudPrueba.tipoSolicitud = "PC"

    Dim rsMapeo As DAO.Recordset ' Se necesita un mock recordset
    ' ... (aquí iría la creación de un recordset mock, se omite por simplicidad)
    mockMapeoRepository.ConfigureGetMapeoPorTipo rsMapeo

    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler

    ' Act
    Dim rutaResultado As String
    rutaResultado = docService.GenerarDocumento(solicitudPrueba)

    ' Assert
    ModAssert.AssertNotEquals "", rutaResultado, "La ruta del documento no debe ser vacía."
    ModAssert.AssertTrue mockWordManager.AbrirDocumento_WasCalled, "Se debió llamar a AbrirDocumento."
    ModAssert.AssertTrue mockWordManager.GuardarDocumento_WasCalled, "Se debió llamar a GuardarDocumento."

    Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente.Pass
    GoTo Cleanup

TestFail:
    Test_GenerarDocumento_ConDatosValidos_DebeGenerarDocumentoCorrectamente.Fail "Error inesperado: " & Err.Description

Cleanup:
    Call Teardown ' Limpia y resetea todos los mocks
End Function

Private Function Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia() As CTestResult
    Set Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia = New CTestResult
    Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia.Initialize "GenerarDocumento con plantilla inexistente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Call Setup
    
    ' Configurar mocks para simular error
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "IS_TEST_MODE", False ' Para que verifique existencia
    Call mockWordManager.ConfigureAbrirDocumento(False) ' Simular fallo
    
    ' Crear solicitud de prueba
    Dim solicitud As ESolicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Inexistente"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' Act
    Dim Resultado As String
    Resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    ModAssert.AssertEquals "", Resultado, "Debería haber retornado cadena vacía"
    
    Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia.Pass
    GoTo Cleanup
    
TestFail:
    Test_GenerarDocumento_ConPlantillaInexistente_DebeRetornarCadenaVacia.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
End Function

Private Function Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia() As CTestResult
    Set Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia = New CTestResult
    Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia.Initialize "GenerarDocumento con error en WordManager"
    
    On Error GoTo TestFail
    
    ' Arrange
    Call Setup
    
    ' Configurar mocks
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "IS_TEST_MODE", True
    Call mockWordManager.ConfigureAbrirDocumento(True)
    Call mockWordManager.ConfigureGuardarDocumento(False) ' Simular fallo al guardar
    
    ' Crear solicitud de prueba
    Dim solicitud As ESolicitud
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Permiso"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' Act
    Dim Resultado As String
    Resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    ModAssert.AssertEquals "", Resultado, "Debería haber retornado cadena vacía"
    
    Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia.Pass
    GoTo Cleanup
    
TestFail:
    Test_GenerarDocumento_ConErrorEnWordManager_DebeRetornarCadenaVacia.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
End Function

' ============================================================================
' PRUEBAS DE LECTURA DE DOCUMENTOS
' ============================================================================

Private Function Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud() As CTestResult
    Set Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud = New CTestResult
    Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud.Initialize "LeerDocumento con documento válido"
    
    On Error GoTo TestFail
    
    ' Arrange
    Call Setup
    
    ' Configurar mocks
    Dim solicitudMock As New ESolicitud
    solicitudMock.idSolicitud = 123
    solicitudMock.tipoSolicitud = "Permiso"
    
    mockRepository.ConfigureGetSolicitudById solicitudMock
    mockRepository.ConfigureSaveSolicitud 1
    Call mockWordManager.ConfigureLeerDocumento("[nombre]Juan Pérez[/nombre][fecha]2025-01-01[/fecha]")
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' Act
    Dim Resultado As Boolean
    Resultado = docService.LeerDocumento("C:\test.docx", 123)
    
    ' Assert
    ModAssert.AssertTrue Resultado, "Fallo al leer documento"
    ModAssert.AssertEquals 1, mockRepository.GetSolicitudById_CallCount, "Se debe llamar a GetSolicitudById exactamente una vez."
    ModAssert.AssertEquals 1, mockRepository.SaveSolicitud_CallCount, "Se debe llamar a SaveSolicitud exactamente una vez."
    
    Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud.Pass
    GoTo Cleanup
    
TestFail:
    Test_LeerDocumento_ConDocumentoValido_DebeActualizarSolicitud.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
End Function

Private Function Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse() As CTestResult
    Set Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse = New CTestResult
    Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse.Initialize "LeerDocumento con documento inexistente"
    
    On Error GoTo TestFail
    
    ' Arrange
    Call Setup
    
    ' Configurar mocks para simular error
    Dim solicitudMock As New ESolicitud
    solicitudMock.idSolicitud = 123
    
    Call mockRepository.ConfigureGetSolicitudById(solicitudMock)
    Call mockWordManager.ConfigureLeerDocumento("") ' Simular fallo
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' Act
    Dim Resultado As Boolean
    Resultado = docService.LeerDocumento("C:\inexistente.docx", 123)
    
    ' Assert
    ModAssert.AssertFalse Resultado, "Debería haber retornado False"
    
    Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse.Pass
    GoTo Cleanup
    
TestFail:
    Test_LeerDocumento_ConDocumentoInexistente_DebeRetornarFalse.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
End Function

' ============================================================================
' PRUEBAS DE INICIALIZACIÓN
' ============================================================================

Private Function Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente() As CTestResult
    Set Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente = New CTestResult
    Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente.Initialize "Initialize con dependencias válidas"
    
    On Error GoTo TestFail
    
    ' Arrange
    Call Setup
    
    ' Act
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' Assert
    ' Verificar que no se produzcan errores durante la inicialización
    Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente.Pass
    GoTo Cleanup
    
TestFail:
    Test_Initialize_ConDependenciasValidas_DebeInicializarCorrectamente.Fail "Error inesperado durante inicialización: " & Err.Description
    
Cleanup:
    Call Teardown
End Function


