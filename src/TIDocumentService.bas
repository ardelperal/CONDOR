Attribute VB_Name = "TIDocumentService"
Option Compare Database
Option Explicit


' =====================================================
' Módulo: IntegrationTestDocumentService
' Descripción: Pruebas de integración para CDocumentService
' Versión: 3.0 (Refactorización completa)
' =====================================================


' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

' --- Variables eliminadas - ahora se declaran localmente en cada función ---

' =====================================================
' SUITE SETUP - PREPARACIÓN INICIAL DE LA SUITE
' =====================================================
Private Sub SuiteSetup()
    ' Asegurarse de que el entorno está limpio antes de empezar
    Call SuiteTeardown

    On Error GoTo ErrorHandler
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()

    ' 1. Crear estructura de directorios
    fs.CreateFolder modTestUtils.GetWorkspacePath() & "doc_service_test\"
    fs.CreateFolder modTestUtils.GetWorkspacePath() & "doc_service_test\templates\"
    fs.CreateFolder modTestUtils.GetWorkspacePath() & "doc_service_test\generated\"

    ' 2. Aprovisionar BD de prueba
    Dim templateDbName As String: templateDbName = "Document_test_template.accdb"
    Dim activeDbName As String: activeDbName = "Document_integration_test.accdb"
    modTestUtils.PrepareTestDatabase templateDbName, activeDbName

    ' 3. Aprovisionar plantilla Word
    Dim plantillaSrc As String
    plantillaSrc = modTestUtils.JoinPath(projectPath, "back\recursos\Plantillas\PC.docx")
    Call modTestUtils.EnsureFolder(modTestUtils.GetWorkspacePath() & "doc_service_test\templates\")
    fs.CopyFile plantillaSrc, modTestUtils.GetWorkspacePath() & "doc_service_test\templates\PC.docx"

    ' 4. Insertar datos maestros en la BD de prueba
    Dim db As DAO.Database
    Dim activePath As String: activePath = modTestUtils.GetWorkspacePath() & activeDbName
    Set db = DBEngine.OpenDatabase(activePath)

    ' BLINDAJE: Limpiar tablas antes de insertar, respetando integridad referencial
    db.Execute "DELETE * FROM tbDatosPC", dbFailOnError
    db.Execute "DELETE * FROM tbMapeoCampos", dbFailOnError
    db.Execute "DELETE * FROM tbSolicitudes", dbFailOnError

    ' Insertar los datos de prueba
    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, tipoSolicitud, codigoSolicitud, idExpediente, idEstadoInterno, usuarioCreacion, fechaCreacion) VALUES (999, 'PC', 'TEST-001', 1, 1, 'test_user', Now())", dbFailOnError
    db.Execute "INSERT INTO tbDatosPC (idSolicitud, refContratoInspeccionOficial) VALUES (999, 'DATO_PRUEBA_CONTRATO')", dbFailOnError
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'refContratoInspeccionOficial', 'MARCADOR_CONTRATO')", dbFailOnError
    
    db.Close

    Set db = Nothing
    Set fs = Nothing
    Exit Sub
ErrorHandler:
    If Not db Is Nothing Then db.Close
    Err.Raise Err.Number, "TIDocumentService.SuiteSetup", Err.Description
End Sub
' =====================================================
' FUNCIÓN PRINCIPAL DEL FRAMEWORK ESTÁNDAR
' =====================================================
Public Function TIDocumentServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIDocumentService (Optimizada)"
    
    On Error GoTo CleanupSuite

    ' 1. Configurar el entorno UNA SOLA VEZ para toda la suite
    Call SuiteSetup
    
    ' 2. Ejecutar todas las pruebas individuales
    suiteResult.AddResult TestGenerarDocumentoSuccess()
    
CleanupSuite:
    ' 3. Limpiar el entorno UNA SOLA VEZ al final, incluso si hay errores
    Call SuiteTeardown
    
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIDocumentServiceRunAll = suiteResult
End Function

' =====================================================
' FUNCIONES SETUP Y TEARDOWN ELIMINADAS
' Cada función de prueba maneja sus propios recursos
' =====================================================

' =====================================================
' TEST DE INTEGRACIÓN PRINCIPAL
' =====================================================
Private Function TestGenerarDocumentoSuccess() As CTestResult
    Set TestGenerarDocumentoSuccess = New CTestResult
    TestGenerarDocumentoSuccess.Initialize "GenerarDocumento debe crear un archivo Word con datos reales"

    ' --- Declarar TODAS las variables de objeto aquí ---
    Dim solicitudService As ISolicitudService
    Dim mapeoRepo As IMapeoRepository
    Dim wordManager As IWordManager
    Dim operationLogger As IOperationLogger
    Dim ErrorHandler As IErrorHandlerService
    Dim documentService As IDocumentService
    Dim fileSystem As IFileSystem
    Dim solicitudPrueba As ESolicitud

    On Error GoTo TestFail

    ' ARRANGE: El entorno ya está creado por SuiteSetup. Solo inicializamos las dependencias.
    InitializeRealDependencies solicitudService, mapeoRepo, wordManager, operationLogger, ErrorHandler, documentService, fileSystem
    
    ' ARRANGE (continuación)
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = modTestUtils.GetWorkspacePath() & "doc_service_test\templates\PC.docx"
    
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    If Not fileSystem.FileExists(templatePath) Then
        Err.Raise vbObjectError + 101, "Test.Arrange", "La plantilla de Word no existe en la ruta esperada: " & templatePath
    End If

    Set solicitudPrueba = solicitudService.ObtenerSolicitudPorId(999)
    modAssert.AssertNotNull solicitudPrueba, "La solicitud de prueba no se pudo cargar desde la BD."
    
    DBEngine.BeginTrans
    
    ' ACT: Ejecutar el método principal a probar
    Dim rutaGenerada As String
    rutaGenerada = documentService.GenerarDocumento(solicitudPrueba)

    ' ASSERT: Verificar los resultados
    modAssert.AssertNotEquals "", rutaGenerada, "La ruta del documento generado no debe estar vacía."
    modAssert.AssertTrue fileSystem.FileExists(rutaGenerada), "El archivo generado debe existir en el disco."
    
    TestGenerarDocumentoSuccess.Pass
    GoTo Cleanup

TestFail:
    TestGenerarDocumentoSuccess.Fail "Error en tiempo de ejecución: " & Err.Description & " en línea " & Erl

Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    
    ' --- LIMPIEZA CRÍTICA DE RECURSOS ---
    
    ' 1. CERRAR WORD PARA EVITAR PROCESOS ZOMBIE
    If Not wordManager Is Nothing Then
        wordManager.Dispose
    End If
    
    ' 2. Liberar todas las variables de objeto
    Set solicitudService = Nothing
    Set mapeoRepo = Nothing
    Set wordManager = Nothing
    Set operationLogger = Nothing
    Set ErrorHandler = Nothing
    Set documentService = Nothing
    Set fileSystem = Nothing
    Set solicitudPrueba = Nothing
End Function

' =====================================================
' SUITE TEARDOWN - LIMPIEZA FINAL DE LA SUITE
' =====================================================
Private Sub SuiteTeardown()
    On Error Resume Next
    Call modTestUtils.CloseAllWordInstancesForTesting
    ' Usar función centralizada de limpieza
    modTestUtils.CleanupTestFolder "doc_service_test\"
End Sub

' =====================================================
' MÉTODOS AUXILIARES PRIVADOS
' =====================================================
Private Sub InitializeRealDependencies(ByRef solicitudService As ISolicitudService, _
                                       ByRef mapeoRepo As IMapeoRepository, ByRef wordManager As IWordManager, _
                                       ByRef operationLogger As IOperationLogger, ByRef ErrorHandler As IErrorHandlerService, _
                                       ByRef documentService As IDocumentService, ByRef fileSystem As IFileSystem)
    
    ' ARRANGE: Crear configuración local apuntando a la BD de prueba de esta suite
    Dim config As IConfig
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "CONDOR_DATA_PATH", modTestUtils.GetWorkspacePath() & "Document_integration_test.accdb"
    mockConfigImpl.SetSetting "CONDOR_PASSWORD", ""
    Set config = mockConfigImpl
    
    ' Crear dependencias inyectando la configuración local
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(config)
    Set wordManager = modWordManagerFactory.CreateWordManager()
    Set solicitudService = modSolicitudServiceFactory.CreateSolicitudService(config)
    Set mapeoRepo = modRepositoryFactory.CreateMapeoRepository(config)
    Set documentService = modDocumentServiceFactory.CreateDocumentService(config)
End Sub






