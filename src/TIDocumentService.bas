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
    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, tipoSolicitud, codigoSolicitud, idExpediente) VALUES (999, 'PC', 'TEST-001', 1)", dbFailOnError
    db.Execute "INSERT INTO tbDatosPC (idSolicitud, refContratoInspeccionOficial) VALUES (999, 'DATO_PRUEBA_CONTRATO')", dbFailOnError
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'refContratoInspeccionOficial', 'MARCADOR_CONTRATO')", dbFailOnError
    db.Close

    Set db = Nothing
    Set fs = Nothing
    Exit Sub
ErrorHandler:
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
    Dim config As IConfig
    Dim solicitudService As ISolicitudService
    Dim mapeoRepo As IMapeoRepository
    Dim wordManager As IWordManager
    Dim operationLogger As IOperationLogger
    Dim errorHandler As IErrorHandlerService
    Dim documentService As IDocumentService
    Dim fileSystem As IFileSystem
    Dim solicitudPrueba As ESolicitud

    On Error GoTo TestFail

    ' ARRANGE: El entorno ya está creado por SuiteSetup. Solo inicializamos las dependencias.
    InitializeRealDependencies config, solicitudService, mapeoRepo, wordManager, operationLogger, errorHandler, documentService, fileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem() ' Necesitamos una instancia para el Assert

    ' Obtener la solicitud de prueba (ID 999) que insertó SuiteSetup
    Set solicitudPrueba = solicitudService.ObtenerSolicitudPorId(999)
    modAssert.AssertNotNull solicitudPrueba, "La solicitud de prueba no se pudo cargar desde la BD."
    
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
    ' --- LIMPIEZA CRÍTICA DE RECURSOS ---
    On Error Resume Next
    
    ' 1. CERRAR WORD PARA EVITAR PROCESOS ZOMBIE
    If Not wordManager Is Nothing Then
        wordManager.Dispose
    End If
    
    ' 2. Liberar todas las variables de objeto
    Set config = Nothing
    Set solicitudService = Nothing
    Set mapeoRepo = Nothing
    Set wordManager = Nothing
    Set operationLogger = Nothing
    Set errorHandler = Nothing
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
Private Sub InitializeRealDependencies(ByRef config As IConfig, ByRef solicitudService As ISolicitudService, _
                                       ByRef mapeoRepo As IMapeoRepository, ByRef wordManager As IWordManager, _
                                       ByRef operationLogger As IOperationLogger, ByRef errorHandler As IErrorHandlerService, _
                                       ByRef documentService As IDocumentService, ByRef fileSystem As IFileSystem)
    
    ' Utilizar el singleton de configuración para coherencia
    Set config = modTestContext.GetTestConfig()
    
    ' Propagar la configuración de prueba a todas las factorías
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config)
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(config)
    Set wordManager = modWordManagerFactory.CreateWordManager(config)
    Set solicitudService = modSolicitudServiceFactory.CreateSolicitudService(config)
    Set mapeoRepo = modRepositoryFactory.CreateMapeoRepository(config)
    Set documentService = modDocumentServiceFactory.CreateDocumentService(config)
End Sub



