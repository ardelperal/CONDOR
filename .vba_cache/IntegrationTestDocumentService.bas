Attribute VB_Name = "IntegrationTest_DocumentService"
' =====================================================
' Módulo: IntegrationTest_DocumentService
' Descripción: Pruebas de integración para CDocumentService
' Versión: 3.0 (Refactorización completa)
' =====================================================

Option Explicit

#If DEV_MODE Then

' --- Constantes para el entorno de prueba ---
Private Const TEST_ENV_PATH As String = "back\test_db\active\doc_service_test\"
Private Const TEST_TEMPLATES_PATH As String = TEST_ENV_PATH & "templates\"
Private Const TEST_GENERATED_PATH As String = TEST_ENV_PATH & "generated\"
Private Const TEST_DB_ACTIVE_PATH As String = TEST_ENV_PATH & "CONDOR_integration_test.accdb"
Private Const SOURCE_TEMPLATE_FILE As String = "back\recursos\Plantillas\PC.docx"
Private Const DB_TEMPLATE_FILE As String = "back\test_db\templates\CONDOR_test_template.accdb"

' --- Variables a nivel de módulo para las dependencias ---
Private m_Config As IConfig
Private m_SolicitudRepo As ISolicitudRepository
Private m_MapeoRepo As IMapeoRepository
Private m_WordManager As IWordManager
Private m_OperationLogger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService
Private m_DocumentService As IDocumentService
Private m_FileSystem As IFileSystem
Private m_ExpedienteRepo As IExpedienteRepository ' Necesario para inicializar el servicio completo

' =====================================================
' FUNCIÓN PRINCIPAL DEL FRAMEWORK ESTÁNDAR
' =====================================================
Public Function IntegrationTest_DocumentService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_DocumentService"

    suiteResult.AddTestResult Test_GenerarDocumento_Success()

    Set IntegrationTest_DocumentService_RunAll = suiteResult
End Function

' =====================================================
' SETUP Y TEARDOWN
' =====================================================
Private Sub Setup()
    On Error GoTo TestError

    Set m_FileSystem = modFileSystemFactory.CreateFileSystem()

    ' 1. Crear estructura de directorios de prueba
    m_FileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_ENV_PATH
    m_FileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH
    m_FileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_GENERATED_PATH

    ' 2. Aprovisionar BD de prueba
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & DB_TEMPLATE_FILE, modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH

    ' 3. Aprovisionar plantilla de Word de prueba
    m_FileSystem.CopyFile modTestUtils.GetProjectPath() & SOURCE_TEMPLATE_FILE, modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH & "PC.docx"

    ' 4. Insertar datos necesarios en la BD de prueba
    InsertTestData

    ' 5. Inicializar todas las dependencias en el orden correcto
    InitializeRealDependencies

    Exit Sub
TestError:
    Err.Raise Err.Number, "IntegrationTest_DocumentService.Setup", "Fallo en el Setup: " & Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next ' Ignorar errores en la limpieza

    ' Cerrar Word si se quedó abierto
    If Not m_WordManager Is Nothing Then m_WordManager.CerrarDocumento

    ' Eliminar directorio de prueba y todo su contenido
    If m_FileSystem Is Nothing Then Set m_FileSystem = modFileSystemFactory.CreateFileSystem()
    Dim fullTestEnvPath As String
    fullTestEnvPath = modTestUtils.GetProjectPath() & TEST_ENV_PATH

    If m_FileSystem.FolderExists(fullTestEnvPath) Then
        m_FileSystem.DeleteFolder fullTestEnvPath
    End If

    ' Liberar todos los objetos
    Set m_Config = Nothing
    Set m_SolicitudRepo = Nothing
    Set m_MapeoRepo = Nothing
    Set m_WordManager = Nothing
    Set m_OperationLogger = Nothing
    Set m_ErrorHandler = Nothing
    Set m_DocumentService = Nothing
    Set m_FileSystem = Nothing
    Set m_ExpedienteRepo = Nothing
End Sub

' =====================================================
' TEST DE INTEGRACIÓN PRINCIPAL
' =====================================================
Private Function Test_GenerarDocumento_Success() As CTestResult
    Set Test_GenerarDocumento_Success = New CTestResult
    Test_GenerarDocumento_Success.Initialize "GenerarDocumento debe crear un archivo Word con datos reales"

    On Error GoTo TestFail

    ' ARRANGE: El Setup ya ha preparado todo el entorno
    Call Setup

    ' Obtener la solicitud de prueba (ID 999) que hemos insertado
    Dim solicitudPrueba As E_Solicitud
    Set solicitudPrueba = m_SolicitudRepo.GetSolicitudById(999)
    modAssert.AssertNotNull solicitudPrueba, "La solicitud de prueba no se pudo cargar desde la BD."

    ' ACT: Ejecutar el método principal a probar
    Dim rutaGenerada As String
    rutaGenerada = m_DocumentService.GenerarDocumento(solicitudPrueba)

    ' ASSERT: Verificar los resultados
    modAssert.AssertNotEquals "", rutaGenerada, "La ruta del documento generado no debe estar vacía."
    modAssert.AssertTrue m_FileSystem.FileExists(rutaGenerada), "El archivo generado debe existir en el disco."

    Test_GenerarDocumento_Success.Pass
    GoTo Cleanup

TestFail:
    Test_GenerarDocumento_Success.Fail "Error en tiempo de ejecución: " & Err.Description & " en línea " & Erl

Cleanup:
    Call Teardown
End Function

' =====================================================
' MÉTODOS AUXILIARES PRIVADOS
' =====================================================
Private Sub InitializeRealDependencies()
    ' Crea e inicializa todas las dependencias en el orden correcto

    ' 1. Crear configuración de prueba
    Dim testConfig As New CConfig
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    settings.Add modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH, "PLANTILLA_PATH"
    settings.Add modTestUtils.GetProjectPath() & TEST_GENERATED_PATH, "GENERATED_DOCS_PATH"
    testConfig.LoadFromCollection settings
    Set m_Config = testConfig

    Set m_ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(m_Config, m_FileSystem)

    ' 2. Repositorios usando factory con testConfig
    Set m_MapeoRepo = modRepositoryFactory.CreateMapeoRepository(testConfig, m_ErrorHandler)
    Set m_SolicitudRepo = modRepositoryFactory.CreateSolicitudRepository(testConfig, m_ErrorHandler)
    Set m_ExpedienteRepo = modRepositoryFactory.CreateExpedienteRepository(testConfig, m_ErrorHandler)

    ' 3. Servicios de Infraestructura
    Set m_OperationLogger = modOperationLoggerFactory.CreateOperationLogger()
    Set m_WordManager = modWordManagerFactory.CreateWordManager()

    ' 4. Servicio Principal a Probar usando factory con testConfig
    Set m_DocumentService = modDocumentServiceFactory.CreateDocumentService(testConfig)
End Sub

Private Sub InsertTestData()
    ' Inserta el mínimo de datos necesarios en la BD de prueba activa
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & TEST_DB_ACTIVE_PATH)

    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, tipoSolicitud, codigoSolicitud, idExpediente) VALUES (999, 'PC', 'TEST-001', 1)"
    db.Execute "INSERT INTO T_Datos_PC (idSolicitud, Parte0_1) VALUES (999, 'DATO_PRUEBA_PARTE0_1')"
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'Parte0_1', 'MARCADOR_PARTE0_1')"

    db.Close
    Set db = Nothing
End Sub

#End If