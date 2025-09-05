Attribute VB_Name = "TIDocumentService"
Option Compare Database
Option Explicit


' =====================================================
' Módulo: IntegrationTestDocumentService
' Descripción: Pruebas de integración para CDocumentService
' Versión: 3.0 (Refactorización completa)
' =====================================================


' --- Constantes para el entorno de prueba ---
Private Const TEST_ENV_PATH As String = "back\test_db\active\doc_service_test\"
Private Const TEST_TEMPLATES_PATH As String = TEST_ENV_PATH & "templates\"
Private Const TEST_GENERATED_PATH As String = TEST_ENV_PATH & "generated\"
Private Const TEST_DB_ACTIVE_PATH As String = TEST_ENV_PATH & "CONDOR_integration_test.accdb"
Private Const SOURCE_TEMPLATE_FILE As String = "back\recursos\Plantillas\PC.docx"
Private Const DB_TEMPLATE_FILE As String = "back\test_db\templates\CONDOR_test_template.accdb"

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
    fs.CreateFolder projectPath & TEST_ENV_PATH
    fs.CreateFolder projectPath & TEST_TEMPLATES_PATH
    fs.CreateFolder projectPath & TEST_GENERATED_PATH

    ' 2. Aprovisionar BD de prueba
    modTestUtils.PrepareTestDatabase projectPath & DB_TEMPLATE_FILE, projectPath & TEST_DB_ACTIVE_PATH

    ' 3. Aprovisionar plantilla Word
    fs.CopyFile projectPath & SOURCE_TEMPLATE_FILE, projectPath & TEST_TEMPLATES_PATH & "PC.docx"

    ' 4. Insertar datos maestros en la BD de prueba
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(projectPath & TEST_DB_ACTIVE_PATH)

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
    Dim templatePath As String: templatePath = projectPath & TEST_TEMPLATES_PATH & "PC.docx"
    
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
    On Error Resume Next ' Blindaje
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFolderRecursive modTestUtils.GetProjectPath() & TEST_ENV_PATH
    Set fs = Nothing
End Sub

' =====================================================
' MÉTODOS AUXILIARES PRIVADOS
' =====================================================
Private Sub InitializeRealDependencies(ByRef solicitudService As ISolicitudService, _
                                       ByRef mapeoRepo As IMapeoRepository, ByRef wordManager As IWordManager, _
                                       ByRef operationLogger As IOperationLogger, ByRef ErrorHandler As IErrorHandlerService, _
                                       ByRef documentService As IDocumentService, ByRef fileSystem As IFileSystem)
    
    ' Las factorías obtienen la configuración del contexto centralizado
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    Set wordManager = modWordManagerFactory.CreateWordManager()
    Set solicitudService = modSolicitudServiceFactory.CreateSolicitudService()
    Set mapeoRepo = modRepositoryFactory.CreateMapeoRepository()
    Set documentService = modDocumentServiceFactory.CreateDocumentService()
End Sub






