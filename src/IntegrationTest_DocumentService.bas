Attribute VB_Name = "IntegrationTest_DocumentService"
' =====================================================
' Módulo: IntegrationTest_DocumentService
' Descripción: Pruebas de integración para CDocumentService
' Autor: CONDOR-Expert
' Fecha: 2024
' =====================================================

Option Explicit

' Constantes para el entorno de prueba
Private Const TEST_ENV_PATH As String = "back\test_env\docs"
Private Const TEST_TEMPLATES_PATH As String = "back\test_env\docs\templates"
Private Const TEST_DB_PATH As String = "back\test_env\docs\CONDOR_integration_test.accdb"
Private Const SOURCE_TEMPLATE_PATH As String = "back\recursos\Plantillas\PC.docx"
Private Const TEST_DB_TEMPLATE As String = "back\CONDOR_datos.accdb"

' Variables globales para el entorno de prueba
Private m_Config As CConfig
Private m_SolicitudRepo As CSolicitudRepository
Private m_ExpedienteRepo As CExpedienteRepository
Private m_MapeoRepo As CMapeoRepository
Private m_WordManager As CWordManager
Private m_OperationLogger As COperationLogger
Private m_ErrorHandler As CErrorHandlerService
Private m_DocumentService As CDocumentService
Private m_FileSystem As IFileSystem

' =====================================================
' FUNCIÓN PRINCIPAL DEL FRAMEWORK ESTÁNDAR
' =====================================================
Public Function IntegrationTest_DocumentService_RunAll() As CTestSuiteResult
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_DocumentService"
    
    ' Ejecutar tests de integración
    suiteResult.AddTestResult Test_GenerarDocumento_Integration_Success()
    
    Set IntegrationTest_DocumentService_RunAll = suiteResult
End Function

' =====================================================
' SETUP - PREPARACIÓN DEL ENTORNO DE PRUEBA
' =====================================================
Private Sub Setup()
    ' Inicializar IFileSystem usando factory
    Set m_FileSystem = modFileSystemFactory.CreateFileSystem()
    
    ' Crear estructura de directorios de prueba
    CreateTestDirectories
    
    ' Copiar plantilla de Word
    CopyTestTemplate
    
    ' Crear base de datos de prueba
    CreateTestDatabase
    
    ' Inicializar dependencias reales
    InitializeRealDependencies
End Sub

' =====================================================
' TEARDOWN - LIMPIEZA DEL ENTORNO DE PRUEBA
' =====================================================
Private Sub Teardown()
    ' Cerrar conexiones y liberar objetos
    Set m_DocumentService = Nothing
    Set m_Config = Nothing
    Set m_SolicitudRepo = Nothing
    Set m_ExpedienteRepo = Nothing
    Set m_MapeoRepo = Nothing
    Set m_WordManager = Nothing
    Set m_OperationLogger = Nothing
    Set m_ErrorHandler = Nothing
    
    ' Eliminar directorio de prueba y todo su contenido usando IFileSystem
    Dim testEnvPath As String
    testEnvPath = modTestUtils.GetProjectPath() & TEST_ENV_PATH
    
    If m_FileSystem.FolderExists(testEnvPath) Then
        m_FileSystem.DeleteFolder testEnvPath
    End If
    
    Set m_FileSystem = Nothing
End Sub

' =====================================================
' TESTS DE INTEGRACIÓN
' =====================================================
Private Function Test_GenerarDocumento_Integration_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_GenerarDocumento_Integration_Success"
    
    On Error GoTo TestError
    
    ' Arrange
    Setup
    
    ' Obtener solicitud de prueba desde el repositorio
    Dim solicitudPrueba As T_Solicitud
    Set solicitudPrueba = GetTestSolicitud()
    
    ' Act
    Dim rutaGenerada As String
    rutaGenerada = m_DocumentService.GenerarDocumento(solicitudPrueba)
    
    ' Assert
    modAssert.AssertNotEmpty rutaGenerada, "La ruta del documento generado no debe estar vacía"
    modAssert.AssertTrue m_FileSystem.FileExists(rutaGenerada), "El archivo generado debe existir en el disco"
    
    ' Verificar contenido del documento generado
    VerifyDocumentContent rutaGenerada, solicitudPrueba
    
    testResult.SetPassed
    
TestCleanup:
    ' Cleanup
    Teardown
    Set Test_GenerarDocumento_Integration_Success = testResult
    Exit Function
    
TestError:
    testResult.SetFailed "Error en Test_GenerarDocumento_Integration_Success: " & Err.Description
    Resume TestCleanup
End Function

' =====================================================
' MÉTODOS AUXILIARES PRIVADOS
' =====================================================
Private Sub CreateTestDirectories()
    ' Crear directorio principal de prueba
    If Not m_FileSystem.FolderExists(modTestUtils.GetProjectPath() & TEST_ENV_PATH) Then
        m_FileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_ENV_PATH
    End If
    
    ' Crear subdirectorio de plantillas
    If Not m_FileSystem.FolderExists(modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH) Then
        m_FileSystem.CreateFolder modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH
    End If
End Sub

Private Sub CopyTestTemplate()
    ' Verificar que la plantilla fuente existe
    If Not m_FileSystem.FileExists(modTestUtils.GetProjectPath() & SOURCE_TEMPLATE_PATH) Then
        Err.Raise 53, "IntegrationTest_DocumentService", "No se encontró la plantilla PC.docx en: " & modTestUtils.GetProjectPath() & SOURCE_TEMPLATE_PATH
    End If
    
    ' Copiar plantilla al directorio de prueba
    Dim destPath As String
    destPath = modTestUtils.GetProjectPath() & TEST_TEMPLATES_PATH & "\PC.docx"
    m_FileSystem.CopyFile modTestUtils.GetProjectPath() & SOURCE_TEMPLATE_PATH, destPath
End Sub

Private Sub CreateTestDatabase()
    ' Copiar base de datos plantilla
    If Not m_FileSystem.FileExists(modTestUtils.GetProjectPath() & TEST_DB_TEMPLATE) Then
        Err.Raise 53, "IntegrationTest_DocumentService", "No se encontró la BD plantilla en: " & modTestUtils.GetProjectPath() & TEST_DB_TEMPLATE
    End If
    
    m_FileSystem.CopyFile modTestUtils.GetProjectPath() & TEST_DB_TEMPLATE, modTestUtils.GetProjectPath() & TEST_DB_PATH
    
    ' Insertar datos de prueba
    InsertTestData
End Sub

Private Sub InsertTestData()
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Conectar a la base de datos de prueba
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & modTestUtils.GetProjectPath() & TEST_DB_PATH & ";"
    
    ' Insertar solicitud de prueba
    Dim sqlSolicitud As String
    sqlSolicitud = "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, TipoSolicitud, EstadoInterno, FechaCreacion) " & _
                   "VALUES (999, 1, 'PC', 'BORRADOR', #" & Date & "#)"
    conn.Execute sqlSolicitud
    
    ' Insertar datos específicos PC
    Dim sqlDatosPC As String
    sqlDatosPC = "INSERT INTO T_Datos_PC (idSolicitud, Parte0_1, Parte0_2, Parte0_3) " & _
                 "VALUES (999, 'DATO_PRUEBA_PARTE0_1', 'DATO_PRUEBA_PARTE0_2', 'DATO_PRUEBA_PARTE0_3')"
    conn.Execute sqlDatosPC
    
    ' Insertar mapeos de campos
    Dim sqlMapeo As String
    sqlMapeo = "INSERT INTO tbMapeoCampos (TipoSolicitud, CampoPlantilla, CampoBaseDatos, TablaOrigen) " & _
               "VALUES ('PC', 'Parte0_1', 'Parte0_1', 'T_Datos_PC')"
    conn.Execute sqlMapeo
    
    conn.Close
    Set conn = Nothing
End Sub

Private Sub InitializeRealDependencies()
    ' Configurar la ruta de la base de datos para las pruebas
    Dim settings As Object
    Set settings = CreateObject("Scripting.Dictionary")
    settings.Add "DATABASE_PATH", modTestUtils.GetProjectPath() & TEST_DB_PATH
    settings.Add "DB_PASSWORD", ""
    
    ' Inicializar configuración real
    Set m_Config = New CConfig
    m_Config.Initialize settings
    
    ' Inicializar repositorios reales
    Set m_ExpedienteRepo = New CExpedienteRepository
    m_ExpedienteRepo.Initialize m_Config
    
    Set m_SolicitudRepo = New CSolicitudRepository
    m_SolicitudRepo.Initialize m_Config
    
    ' Inicializar servicio real
    Set m_DocumentService = New CDocumentService
    m_DocumentService.Initialize m_Config, m_ExpedienteRepo, m_SolicitudRepo
End Sub

Private Function GetTestSolicitud() As T_Solicitud
    ' Obtener la solicitud de prueba desde el repositorio
    Set GetTestSolicitud = m_SolicitudRepo.GetById(999)
End Function

Private Sub VerifyDocumentContent(rutaDocumento As String, solicitud As T_Solicitud)
    ' Abrir el documento generado con Word
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Dim doc As Object
    Set doc = wordApp.Documents.Open(rutaDocumento)
    
    ' Leer el contenido del documento
    Dim contenido As String
    contenido = doc.Content.Text
    
    ' Verificar que el marcador Parte0_1 ha sido reemplazado
    modAssert.AssertTrue InStr(contenido, "DATO_PRUEBA_PARTE0_1") > 0, _
                        "El documento debe contener el dato reemplazado 'DATO_PRUEBA_PARTE0_1'"
    
    ' Verificar que no quedan marcadores sin reemplazar
    modAssert.AssertFalse InStr(contenido, "{{Parte0_1}}") > 0, _
                         "El documento no debe contener marcadores sin reemplazar"
    
    ' Cerrar documento y aplicación Word
    doc.Close False
    wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing
End Sub