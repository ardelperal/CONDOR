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

' =====================================================
' PRUEBAS ESPECÍFICAS PARA ExtraerValorMarcador
' =====================================================
Public Function TestExtraerValorMarcadorConMarcadorValidoDebeRetornarValor() As CTestResult
    Set TestExtraerValorMarcadorConMarcadorValidoDebeRetornarValor = New CTestResult
    TestExtraerValorMarcadorConMarcadorValidoDebeRetornarValor.Initialize "ExtraerValorMarcador con marcador válido debe retornar valor correcto"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim rs As DAO.Recordset
    Dim valorExtraido As String
    
    ' Inicializar mocks con validación
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    Set docService = New CDocumentService
    
    ' Validar inicialización de mocks
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' Configurar el servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' ARRANGE: Configurar recordset con datos de mapeo
    Set rs = CreateInMemoryRecordset("MARCADOR_PARTE0_1", "Parte0_1")
    If rs Is Nothing Then GoTo TestFail
    
    ' Validar que el recordset tiene datos
    If rs.EOF And rs.BOF Then
        GoTo TestFail
    End If
    
    ' ACT: Extraer valor del marcador
    valorExtraido = docService.ExtraerValorMarcador("MARCADOR_PARTE0_1", rs)
    
    ' ASSERT: Verificar que se extrajo el valor correcto
    modAssert.AssertEquals "Parte0_1", valorExtraido, "El valor extraído debe corresponder al campo de la tabla"
    
    TestExtraerValorMarcadorConMarcadorValidoDebeRetornarValor.Pass
    GoTo Cleanup
    
TestFail:
    TestExtraerValorMarcadorConMarcadorValidoDebeRetornarValor.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"
    
Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar recordset primero
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
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

Public Function TestExtraerValorMarcadorConMarcadorInexistenteDebeRetornarCadenaVacia() As CTestResult
    Set TestExtraerValorMarcadorConMarcadorInexistenteDebeRetornarCadenaVacia = New CTestResult
    TestExtraerValorMarcadorConMarcadorInexistenteDebeRetornarCadenaVacia.Initialize "ExtraerValorMarcador con marcador inexistente debe retornar cadena vacía"
    
    On Error GoTo TestFail
    
    ' Declarar variables locales
    Dim docService As CDocumentService
    Dim mockConfig As CMockConfig
    Dim mockRepository As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockWordManager As CMockWordManager
    Dim mockMapeoRepository As CMockMapeoRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim rs As DAO.Recordset
    Dim valorExtraido As String
    
    ' Inicializar mocks con validación
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    Set docService = New CDocumentService
    
    ' Validar inicialización de mocks
    If docService Is Nothing Or mockConfig Is Nothing Or mockRepository Is Nothing Or _
       mockLogger Is Nothing Or mockWordManager Is Nothing Or mockMapeoRepository Is Nothing Or _
       mockErrorHandler Is Nothing Then
        GoTo TestFail
    End If
    
    ' Configurar el servicio
    docService.Initialize mockConfig, mockRepository, mockLogger, mockWordManager, mockMapeoRepository, mockErrorHandler
    
    ' ARRANGE: Configurar recordset con datos de mapeo (marcador diferente)
    Set rs = CreateInMemoryRecordset("MARCADOR_PARTE0_1", "Parte0_1")
    If rs Is Nothing Then GoTo TestFail
    
    ' Validar que el recordset tiene datos
    If rs.EOF And rs.BOF Then
        GoTo TestFail
    End If
    
    ' ACT: Intentar extraer valor de un marcador que no existe
    valorExtraido = docService.ExtraerValorMarcador("MARCADOR_INEXISTENTE", rs)
    
    ' ASSERT: Verificar que retorna cadena vacía
    modAssert.AssertEquals "", valorExtraido, "El valor extraído debe ser cadena vacía para marcador inexistente"
    
    TestExtraerValorMarcadorConMarcadorInexistenteDebeRetornarCadenaVacia.Pass
    GoTo Cleanup
    
TestFail:
    TestExtraerValorMarcadorConMarcadorInexistenteDebeRetornarCadenaVacia.Fail "Error inesperado: " & Err.Description & " (Número: " & Err.Number & ")"
    
Cleanup:
    ' Limpieza robusta de recursos
    On Error Resume Next
    
    ' Limpiar recordset primero
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
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
    Dim rsMapeo As DAO.Recordset
    Dim solicitudPrueba As ESolicitud
    Dim rutaResultado As String
    
    ' Inicializar mocks con validación
    Set docService = New CDocumentService
    Set mockConfig = New CMockConfig
    Set mockRepository = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockWordManager = New CMockWordManager
    Set mockMapeoRepository = New CMockMapeoRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    
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

    Set rsMapeo = CreateInMemoryRecordset()
    If rsMapeo Is Nothing Then GoTo TestFail
    
    mockMapeoRepository.ConfigureGetMapeoPorTipo rsMapeo

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
    
    ' Limpiar recordset primero
    If Not rsMapeo Is Nothing Then
        If rsMapeo.State = adStateOpen Then rsMapeo.Close
        Set rsMapeo = Nothing
    End If
    
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
    Set docService = New CDocumentService
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

Private Function CreateInMemoryRecordset(Optional ByVal marcadorWord As String = "MARCADOR_PARTE0_1", Optional ByVal campoTabla As String = "Parte0_1") As DAO.Recordset
    On Error GoTo TestFail
    
    ' Validar parámetros de entrada
    If Len(Trim(marcadorWord)) = 0 Then marcadorWord = "MARCADOR_PARTE0_1"
    If Len(Trim(campoTabla)) = 0 Then campoTabla = "Parte0_1"
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tempDbPath As String
    Dim errorOccurred As Boolean
    
    errorOccurred = False
    
    ' Crear base de datos temporal usando DAO puro
    tempDbPath = Environ("TEMP") & "\TestMapeo_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    
    ' Eliminar archivo temporal si ya existe
    On Error Resume Next
    If Dir(tempDbPath) <> "" Then
        Kill tempDbPath
    End If
    On Error GoTo TestFail
    
    Set db = DBEngine.CreateDatabase(tempDbPath, dbLangGeneral)
    
    ' Crear tabla tbMapeoCampos con la estructura correcta
    db.Execute "CREATE TABLE tbMapeoCampos (" & _
               "NombreCampoTabla TEXT(255), " & _
               "NombreCampoWord TEXT(255), " & _
               "ValorAsociado TEXT(255))"
    
    ' Insertar datos de prueba necesarios usando parámetros con escape de comillas
    Dim safeCampoTabla As String
    Dim safeMarcadorWord As String
    safeCampoTabla = Replace(campoTabla, "'", "''")
    safeMarcadorWord = Replace(marcadorWord, "'", "''")
    
    db.Execute "INSERT INTO tbMapeoCampos (NombreCampoTabla, NombreCampoWord, ValorAsociado) " & _
               "VALUES ('" & safeCampoTabla & "', '" & safeMarcadorWord & "', '')"
    
    ' Abrir recordset desde la tabla temporal
    Set rs = db.OpenRecordset("SELECT * FROM tbMapeoCampos", dbOpenDynaset)
    
    ' Verificar que el recordset se creó correctamente
    If rs Is Nothing Then
        errorOccurred = True
        GoTo TestFail
    End If
    
    ' Cerrar la base de datos pero mantener el recordset abierto
    ' El recordset mantendrá su propia conexión
    db.Close
    Set db = Nothing
    
    Set CreateInMemoryRecordset = rs
    
    Exit Function
    
TestFail:
    Debug.Print "Error en CreateInMemoryRecordset: " & Err.Description & " (Número: " & Err.Number & ")"
    
    ' Limpieza robusta en caso de error
    On Error Resume Next
    
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    If Not db Is Nothing Then
        If db.Name <> "" Then db.Close
        Set db = Nothing
    End If
    
    ' Eliminar archivo temporal si existe
    If Dir(tempDbPath) <> "" Then
        Kill tempDbPath
    End If
    
    On Error GoTo 0
    
    Set CreateInMemoryRecordset = Nothing
End Function

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
    Set docService = New CDocumentService
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
    Set docService = New CDocumentService
    
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
    Set docService = New CDocumentService
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
    Set docService = New CDocumentService
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


