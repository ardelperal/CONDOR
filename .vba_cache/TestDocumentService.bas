Attribute VB_Name = "TestDocumentService"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: TestDocumentService
' Descripción: Suite de pruebas unitarias para DocumentService
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' Versión: 2.0 - Estandarizado según framework
' ============================================================================

' ============================================================================
' VARIABLES A NIVEL DE MÓDULO - ELIMINADAS PARA MEJORAR ESTABILIDAD
' ============================================================================
' Las variables ahora se declaran localmente en cada función de prueba

' =====================================================
' SETUP Y TEARDOWN
' =====================================================

' Setup y Teardown eliminados - cada función de prueba maneja sus propios recursos

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE DE PRUEBAS
' ============================================================================

Public Function TestDocumentServiceRunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    Call suite.Initialize("DocumentService")
    
    ' Ejecutar todas las pruebas y añadir resultados
    Call suite.AddTestResult(TestGenerarDocumentoConDatosValidosDebeGenerarDocumentoCorrectamente())
    Call suite.AddTestResult(TestGenerarDocumentoConPlantillaInexistenteDebeRetornarCadenaVacia())
    Call suite.AddTestResult(TestGenerarDocumentoConErrorEnWordManagerDebeRetornarCadenaVacia())
    Call suite.AddTestResult(TestLeerDocumentoConDocumentoValidoDebeActualizarSolicitud())
    Call suite.AddTestResult(TestLeerDocumentoConDocumentoInexistenteDebeRetornarFalse())
    Call suite.AddTestResult(TestInitializeConDependenciasValidasDebeInicializarCorrectamente())
    
    Set TestDocumentServiceRunAll = suite
End Function





' ============================================================================
' PRUEBAS DE GENERACIÓN DE DOCUMENTOS
' ============================================================================

Private Function TestGenerarDocumentoConDatosValidosDebeGenerarDocumentoCorrectamente() As CTestResult
    Set TestGenerarDocumentoConDatosValidosDebeGenerarDocumentoCorrectamente = New CTestResult
    TestGenerarDocumentoConDatosValidosDebeGenerarDocumentoCorrectamente.Initialize "GenerarDocumento con datos válidos"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim mapeoMock As EMapeo
    Dim solicitudPrueba As ESolicitud
    Dim rutaResultado As String
    
    ' Inicializar mocks con validación
    Dim docService As IDocumentService
    Set docService = New CMockDocumentService
    Set mockConfig = New CMockConfig
    mockConfig.Reset
    Set mockRepository = New CMockSolicitudRepository
    mockRepository.Reset
    Set mockLogger = New CMockOperationLogger
    mockLogger.Reset
    Set mockWordManager = New CMockWordManager
    mockWordManager.Reset
    Set mockMapeoRepository = New CMockMapeoRepository
    mockMapeoRepository.Reset
    Set mockErrorHandler = New CMockErrorHandlerService
    mockErrorHandler.Reset
    
    ' Validar que todos los mocks se inicializaron correctamente
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' Arrange - Configurar mocks para un escenario de éxito
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "GENERATED_DOCS_PATH", "C:\Generated"
    mockConfig.AddSetting "IS_TEST_MODE", True

    mockWordManager.ConfigureAbrirDocumento True
    mockWordManager.ConfigureReemplazarTexto True
    mockWordManager.ConfigureGuardarDocumento True

    Set solicitudPrueba = New ESolicitud
    If solicitudPrueba Is Nothing Then GoTo TestFail
    
    solicitudPrueba.idSolicitud = 123
    solicitudPrueba.codigoSolicitud = "SOL001"
    solicitudPrueba.tipoSolicitud = "PC"

    ' Crear objeto EMapeo mock
    Dim mapeoMock As New EMapeo
    mapeoMock.idMapeo = 1
    mapeoMock.TipoSolicitud = "PC"
    mapeoMock.PlantillaPath = "C:\Templates\PC_Template.docx"
    mapeoMock.CamposRequeridos = "nombre,fecha,descripcion"
    
    mockMapeoRepository.ConfigureGetMapeoPorTipo mapeoMock

    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler

    ' Act
    rutaResultado = docService.GenerarDocumento(solicitudPrueba)

    ' Assert
    ModAssert.AssertNotEquals "", rutaResultado, "La ruta del documento no debe ser vacía."
    ModAssert.AssertTrue mockWordManager.AbrirDocumento_WasCalled, "Se debió llamar a AbrirDocumento."
    ModAssert.AssertTrue mockWordManager.GuardarDocumento_WasCalled, "Se debió llamar a GuardarDocumento."

    TestGenerarDocumentoConDatosValidosDebeGenerarDocumentoCorrectamente.Pass
    GoTo Cleanup

TestFail:
    TestGenerarDocumentoConDatosValidosDebeGenerarDocumentoCorrectamente.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"

Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar objeto mapeo
    Set mapeoMock = Nothing
    
    ' Limpiar mocks
    If Not mockConfig Is Nothing Then
        mockConfig.Reset
        Set mockConfig = Nothing
    End If
    If Not mockRepository Is Nothing Then
        mockRepository.Reset
        Set mockRepository = Nothing
    End If
    If Not mockLogger Is Nothing Then
        mockLogger.Reset
        Set mockLogger = Nothing
    End If
    If Not mockWordManager Is Nothing Then
        mockWordManager.Reset
        Set mockWordManager = Nothing
    End If
    If Not mockMapeoRepository Is Nothing Then
        mockMapeoRepository.Reset
        Set mockMapeoRepository = Nothing
    End If
    If Not mockErrorHandler Is Nothing Then
        mockErrorHandler.Reset
        Set mockErrorHandler = Nothing
    End If
    
    ' Limpiar objetos principales
    Set docService = Nothing
    Set solicitudPrueba = Nothing
    
    On Error GoTo 0
End Function

Private Function TestGenerarDocumentoConPlantillaInexistenteDebeRetornarCadenaVacia() As CTestResult
    Set TestGenerarDocumentoConPlantillaInexistenteDebeRetornarCadenaVacia = New CTestResult
    TestGenerarDocumentoConPlantillaInexistenteDebeRetornarCadenaVacia.Initialize "GenerarDocumento con plantilla inexistente"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim solicitud As ESolicitud
    Dim Resultado As String
    
    ' Inicializar mocks con validación
    Set docService = New CMockDocumentService
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' Validar inicialización de mocks
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' Arrange - Configurar mocks para simular error
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "IS_TEST_MODE", False ' Para que verifique existencia
    Call mockWordManager.ConfigureAbrirDocumento(False) ' Simular fallo
    
    ' Crear solicitud de prueba
    Set solicitud = New ESolicitud
    If solicitud Is Nothing Then GoTo TestFail
    
    solicitud.idSolicitud = 123
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Inexistente"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' Act
    Resultado = docService.GenerarDocumento(solicitud)
    
    ' Assert
    ModAssert.AssertEquals "", Resultado, "Debería haber retornado cadena vacía"
    
    TestGenerarDocumentoConPlantillaInexistenteDebeRetornarCadenaVacia.Pass
    GoTo Cleanup
    
TestFail:
    TestGenerarDocumentoConPlantillaInexistenteDebeRetornarCadenaVacia.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"
    
Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar mocks
    If Not mockConfig Is Nothing Then
        mockConfig.Reset
        Set mockConfig = Nothing
    End If
    If Not mockRepository Is Nothing Then
        mockRepository.Reset
        Set mockRepository = Nothing
    End If
    If Not mockLogger Is Nothing Then
        mockLogger.Reset
        Set mockLogger = Nothing
    End If
    If Not mockWordManager Is Nothing Then
        mockWordManager.Reset
        Set mockWordManager = Nothing
    End If
    If Not mockMapeoRepository Is Nothing Then
        mockMapeoRepository.Reset
        Set mockMapeoRepository = Nothing
    End If
    If Not mockErrorHandler Is Nothing Then
        mockErrorHandler.Reset
        Set mockErrorHandler = Nothing
    End If
    
    Set docService = Nothing
    Set solicitud = Nothing
    
    On Error GoTo 0
End Function

' ============================================================================
' FUNCIONES AUXILIARES PARA MOCKS
' ============================================================================

' Funciones auxiliares eliminadas - ahora se usan objetos de dominio directamente

Private Function TestGenerarDocumentoConErrorEnWordManagerDebeRetornarCadenaVacia() As CTestResult
    Set TestGenerarDocumentoConErrorEnWordManagerDebeRetornarCadenaVacia = New CTestResult
    TestGenerarDocumentoConErrorEnWordManagerDebeRetornarCadenaVacia.Initialize "GenerarDocumento con error en WordManager debe retornar cadena vacía"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim solicitud As ESolicitud
    Dim resultado As String
    
    ' Inicializar mocks con validación
    Set docService = New CMockDocumentService
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' Validar inicialización de mocks
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' ARRANGE: Configurar mocks para simular fallo al guardar
    mockConfig.AddSetting "PLANTILLA_PATH", "C:\Templates"
    mockConfig.AddSetting "IS_TEST_MODE", True
    Call mockWordManager.ConfigureAbrirDocumento(True)
    Call mockWordManager.ConfigureGuardarDocumento(False) ' Simular fallo al guardar
    
    ' Crear solicitud de prueba
    Set solicitud = New ESolicitud
    If solicitud Is Nothing Then GoTo TestFail
    
    solicitud.idSolicitud = 456
    solicitud.codigoSolicitud = "SOL001"
    solicitud.tipoSolicitud = "Permiso"
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' ACT: Intentar generar documento con error en WordManager
    resultado = docService.GenerarDocumento(solicitud)
    
    ' ASSERT: Verificar que se retorna cadena vacía
    ModAssert.AssertEquals "", resultado, "Debe retornar cadena vacía cuando hay error en WordManager"
    
    TestGenerarDocumentoConErrorEnWordManagerDebeRetornarCadenaVacia.Pass
    GoTo Cleanup
    
TestFail:
    TestGenerarDocumentoConErrorEnWordManagerDebeRetornarCadenaVacia.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"
    
Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar mocks
    If Not mockConfig Is Nothing Then
        mockConfig.Reset
        Set mockConfig = Nothing
    End If
    If Not mockRepository Is Nothing Then
        mockRepository.Reset
        Set mockRepository = Nothing
    End If
    If Not mockLogger Is Nothing Then
        mockLogger.Reset
        Set mockLogger = Nothing
    End If
    If Not mockWordManager Is Nothing Then
        mockWordManager.Reset
        Set mockWordManager = Nothing
    End If
    If Not mockMapeoRepository Is Nothing Then
        mockMapeoRepository.Reset
        Set mockMapeoRepository = Nothing
    End If
    If Not mockErrorHandler Is Nothing Then
        mockErrorHandler.Reset
        Set mockErrorHandler = Nothing
    End If
    
    Set docService = Nothing
    Set solicitud = Nothing
    
    On Error GoTo 0
End Function

' ============================================================================
' PRUEBAS DE LECTURA DE DOCUMENTOS
' ============================================================================

Private Function TestLeerDocumentoConDocumentoValidoDebeActualizarSolicitud() As CTestResult
    Set TestLeerDocumentoConDocumentoValidoDebeActualizarSolicitud = New CTestResult
    TestLeerDocumentoConDocumentoValidoDebeActualizarSolicitud.Initialize "LeerDocumento con documento válido debe actualizar solicitud"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim resultado As Boolean
    
    ' Inicializar mocks con validación
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    Set docService = New CMockDocumentService
    
    ' Validar inicialización de mocks
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' ARRANGE: Configurar mocks para simular éxito
    Dim solicitudMock As New ESolicitud
    solicitudMock.idSolicitud = 456
    solicitudMock.tipoSolicitud = "Permiso"
    
    mockRepository.ConfigureGetSolicitudById solicitudMock
    mockRepository.ConfigureSaveSolicitud 1
    Call mockWordManager.ConfigureLeerDocumento("[nombre]Juan Pérez[/nombre][fecha]2025-01-01[/fecha]")
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' ACT: Leer documento válido
    resultado = docService.LeerDocumento("C:\test.docx", 123)
    
    ' ASSERT: Verificar que se retorna True y se realizan las llamadas esperadas
    ModAssert.AssertTrue resultado, "Debe retornar True cuando el documento se lee correctamente"
    ModAssert.AssertEquals 1, mockRepository.GetSolicitudById_CallCount, "Se debe llamar a GetSolicitudById exactamente una vez."
    ModAssert.AssertEquals 1, mockRepository.SaveSolicitud_CallCount, "Se debe llamar a SaveSolicitud exactamente una vez."
    
    TestLeerDocumentoConDocumentoValidoDebeActualizarSolicitud.Pass
    GoTo Cleanup
    
TestFail:
    TestLeerDocumentoConDocumentoValidoDebeActualizarSolicitud.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"
    
Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar mocks
    If Not mockConfig Is Nothing Then
        mockConfig.Reset
        Set mockConfig = Nothing
    End If
    If Not mockRepository Is Nothing Then
        mockRepository.Reset
        Set mockRepository = Nothing
    End If
    If Not mockLogger Is Nothing Then
        mockLogger.Reset
        Set mockLogger = Nothing
    End If
    If Not mockWordManager Is Nothing Then
        mockWordManager.Reset
        Set mockWordManager = Nothing
    End If
    If Not mockMapeoRepository Is Nothing Then
        mockMapeoRepository.Reset
        Set mockMapeoRepository = Nothing
    End If
    If Not mockErrorHandler Is Nothing Then
        mockErrorHandler.Reset
        Set mockErrorHandler = Nothing
    End If
    
    Set docService = Nothing
    
    On Error GoTo 0
End Function

Private Function TestLeerDocumentoConDocumentoInexistenteDebeRetornarFalse() As CTestResult
    Set TestLeerDocumentoConDocumentoInexistenteDebeRetornarFalse = New CTestResult
    TestLeerDocumentoConDocumentoInexistenteDebeRetornarFalse.Initialize "LeerDocumento con documento inexistente debe retornar False"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim resultado As Boolean
    
    ' Inicializar mocks con validación
    Set docService = New CMockDocumentService
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' Validar inicialización de mocks
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' ARRANGE: Configurar mocks para simular documento inexistente
    Dim solicitudMock As New ESolicitud
    solicitudMock.idSolicitud = 123
    
    Call mockRepository.ConfigureGetSolicitudById(solicitudMock)
    Call mockWordManager.ConfigureLeerDocumento("") ' Simular fallo
    
    ' Inicializar servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' ACT: Intentar leer documento inexistente
    resultado = docService.LeerDocumento("C:\temp\documento_inexistente.docx", 123)
    
    ' ASSERT: Verificar que se retorna False
    ModAssert.AssertFalse resultado, "Debe retornar False cuando el documento no existe"
    
    TestLeerDocumentoConDocumentoInexistenteDebeRetornarFalse.Pass
    GoTo Cleanup
    
TestFail:
    TestLeerDocumentoConDocumentoInexistenteDebeRetornarFalse.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"
    
Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar mocks
    If Not mockConfig Is Nothing Then
        mockConfig.Reset
        Set mockConfig = Nothing
    End If
    If Not mockRepository Is Nothing Then
        mockRepository.Reset
        Set mockRepository = Nothing
    End If
    If Not mockLogger Is Nothing Then
        mockLogger.Reset
        Set mockLogger = Nothing
    End If
    If Not mockWordManager Is Nothing Then
        mockWordManager.Reset
        Set mockWordManager = Nothing
    End If
    If Not mockMapeoRepository Is Nothing Then
        mockMapeoRepository.Reset
        Set mockMapeoRepository = Nothing
    End If
    If Not mockErrorHandler Is Nothing Then
        mockErrorHandler.Reset
        Set mockErrorHandler = Nothing
    End If
    
    Set docService = Nothing
    
    On Error GoTo 0
End Function

' ============================================================================
' PRUEBAS DE INICIALIZACIÓN
' ============================================================================

Private Function TestInitializeConDependenciasValidasDebeInicializarCorrectamente() As CTestResult
    Set TestInitializeConDependenciasValidasDebeInicializarCorrectamente = New CTestResult
    TestInitializeConDependenciasValidasDebeInicializarCorrectamente.Initialize "Initialize con dependencias válidas debe inicializar correctamente"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    
    ' Inicializar mocks con validación
    Set docService = New CMockDocumentService
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' Validar inicialización de mocks
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' ACT: Inicializar con dependencias válidas
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' ASSERT: Verificar que no hay errores (si llegamos aquí, la inicialización fue exitosa)
    TestInitializeConDependenciasValidasDebeInicializarCorrectamente.Pass
    GoTo Cleanup
    
TestFail:
    TestInitializeConDependenciasValidasDebeInicializarCorrectamente.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"
    
Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar mocks
    If Not mockConfig Is Nothing Then
        mockConfig.Reset
        Set mockConfig = Nothing
    End If
    If Not mockRepository Is Nothing Then
        mockRepository.Reset
        Set mockRepository = Nothing
    End If
    If Not mockLogger Is Nothing Then
        mockLogger.Reset
        Set mockLogger = Nothing
    End If
    If Not mockWordManager Is Nothing Then
        mockWordManager.Reset
        Set mockWordManager = Nothing
    End If
    If Not mockMapeoRepository Is Nothing Then
        mockMapeoRepository.Reset
        Set mockMapeoRepository = Nothing
    End If
    If Not mockErrorHandler Is Nothing Then
        mockErrorHandler.Reset
        Set mockErrorHandler = Nothing
    End If
    
    Set docService = Nothing
    
    On Error GoTo 0
End Function


