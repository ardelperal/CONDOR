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
' FUNCIÓN PRINCIPAL DEL FRAMEWORK ESTÁNDAR
' =====================================================
Public Function TIDocumentServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIDocumentService"

    Call suiteResult.AddResult(TestGenerarDocumentoSuccess())

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

    On Error GoTo TestFail

    ' Declarar variables locales
    Dim config As IConfig
    Dim solicitudService As ISolicitudService
    Dim mapeoRepo As IMapeoRepository
    Dim wordManager As IWordManager
    Dim operationLogger As IOperationLogger
    Dim errorHandler As IErrorHandlerService
    Dim documentService As IDocumentService
    Dim fileSystem As IFileSystem
    Dim expedienteRepo As IExpedienteRepository

    ' ARRANGE: Preparar el entorno de prueba
    Set fileSystem = modFileSystemFactory.CreateFileSystem()

    ' 1. Crear estructura de directorios de prueba
    fileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_ENV_PATH
    fileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH
    fileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_GENERATED_PATH

    ' 2. Aprovisionar BD de prueba
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & DB_TEMPLATE_FILE, modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH

    ' 3. Aprovisionar plantilla de Word de prueba
    fileSystem.CopyFile modTestUtils.GetProjectPath() & SOURCE_TEMPLATE_FILE, modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH & "PC.docx"

    ' 4. Insertar datos necesarios en la BD de prueba
    InsertTestData

    ' 5. Inicializar todas las dependencias en el orden correcto
    InitializeRealDependencies config, solicitudService, mapeoRepo, wordManager, operationLogger, errorHandler, documentService, fileSystem, expedienteRepo

    ' Obtener la solicitud de prueba (ID 999) que hemos insertado
    Dim solicitudPrueba As ESolicitud
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
    Call Teardown
    
    ' Liberar todos los objetos
    Set config = Nothing
    Set solicitudService = Nothing
    Set mapeoRepo = Nothing
    Set wordManager = Nothing
    Set operationLogger = Nothing
    Set errorHandler = Nothing
    Set documentService = Nothing
    Set fileSystem = Nothing
    Set expedienteRepo = Nothing
End Function

' =====================================================
' TEARDOWN - LIMPIEZA ROBUSTA DEL ENTORNO DE PRUEBA
' =====================================================
Private Sub Teardown()
    On Error Resume Next ' Blindaje contra cualquier error en la limpieza
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' Usar el nuevo método de borrado recursivo
    fs.DeleteFolderRecursive modTestUtils.GetProjectPath() & TEST_ENV_PATH
    
    Set fs = Nothing
End Sub

' =====================================================
' MÉTODOS AUXILIARES PRIVADOS
' =====================================================
Private Sub InitializeRealDependencies(ByRef config As IConfig, ByRef solicitudService As ISolicitudService, _
                                       ByRef mapeoRepo As IMapeoRepository, ByRef wordManager As IWordManager, _
                                       ByRef operationLogger As IOperationLogger, ByRef errorHandler As IErrorHandlerService, _
                                       ByRef documentService As IDocumentService, ByRef fileSystem As IFileSystem, _
                                       ByRef expedienteRepo As IExpedienteRepository)
    ' Crea e inicializa todas las dependencias en el orden correcto

    ' 1. Crear configuración de prueba
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH
    config.SetSetting "DB_PASSWORD", "" ' La BD de prueba no tiene contraseña
    config.SetSetting "PLANTILLA_PATH", modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH
    config.SetSetting "GENERATED_DOCS_PATH", modTestUtils.GetProjectPath() & TEST_GENERATED_PATH

    ' 2. Crear servicios y repositorios propagando la configuración de prueba
    '    CORRECCIÓN CRÍTICA: Propagar config a TODAS las factorías para evitar configuraciones vacías
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config)
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(config)
    Set wordManager = modWordManagerFactory.CreateWordManager(config)
    Set solicitudService = modSolicitudServiceFactory.CreateSolicitudService(config)
    Set mapeoRepo = modRepositoryFactory.CreateMapeoRepository(config)

    ' 3. Finalmente, crear el servicio principal a probar propagando la configuración
    '    La factoría se encargará de crear y conectar todas las piezas con la configuración correcta
    Set documentService = modDocumentServiceFactory.CreateDocumentService(config)

End Sub

Private Sub InsertTestData()
    ' Inserta el mínimo de datos necesarios en la BD de prueba activa
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH)

    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, tipoSolicitud, codigoSolicitud, idExpediente) VALUES (999, 'PC', 'TEST-001', 1)"
    ' CORRECCIÓN: Usar un campo real de la tabla tbDatosPC
    db.Execute "INSERT INTO tbDatosPC (idSolicitud, refContratoInspeccionOficial) VALUES (999, 'DATO_PRUEBA_CONTRATO')"
    ' CORRECCIÓN: Alinear el mapeo con el campo real
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'refContratoInspeccionOficial', 'MARCADOR_CONTRATO')"

    db.Close
    Set db = Nothing
End Sub



