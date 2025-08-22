Option Compare Database
Option Explicit
' ============================================================================
' Módulo: Test_DocumentService
' Descripción: Pruebas unitarias para CDocumentService.cls, siguiendo la arquitectura TDD.
' ============================================================================

' --- DECLARACIÓN DE MOCKS ---
Private m_MockConfig As CMockConfig
Private m_MockSolicitudRepository As CMockSolicitudRepository
Private m_MockOperationLogger As CMockOperationLogger
Private m_MockSolicitud As CMockSolicitud ' Un mock para el objeto ISolicitud

' --- FUNCIÓN PRINCIPAL DE LA SUITE ---
Public Function Test_DocumentService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.SuiteName = "Test_DocumentService"

    ' Añadir aquí cada una de las pruebas individuales
    suiteResult.AddTest "Test_GenerarDocumento_RutaValida_Success", Test_GenerarDocumento_RutaValida_Success()
    suiteResult.AddTest "Test_GenerarDocumento_PlantillaNoExiste_Fail", Test_GenerarDocumento_PlantillaNoExiste_Fail()
    ' TODO: Añadir más pruebas para LeerDocumento

    Set Test_DocumentService_RunAll = suiteResult
End Function

' --- SETUP DE PRUEBAS ---
Private Sub Setup()
    ' Se ejecuta antes de cada prueba para inicializar los mocks
    Set m_MockConfig = New CMockConfig
    Set m_MockSolicitudRepository = New CMockSolicitudRepository
    Set m_MockOperationLogger = New CMockOperationLogger
    Set m_MockSolicitud = New CMockSolicitud
End Sub

' --- PRUEBAS INDIVIDUALES ---

Private Function Test_GenerarDocumento_RutaValida_Success() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_GenerarDocumento_RutaValida_Success"
    On Error GoTo TestFail

    ' Arrange
    Call Setup
    Dim docService As IDocumentService
    Set docService = CreateTestableDocumentService()

    ' Configurar mocks para un escenario exitoso - SIN dependencias del sistema de ficheros
    m_MockConfig.SetPlantillaPath "C:\Plantillas\PC.docx"
    m_MockConfig.SetGeneratedDocsPath "C:\Documentos\Generados"
    m_MockConfig.IsTestMode = True ' Esto le dice al servicio que no compruebe si el fichero existe
    m_MockSolicitud.tipoSolicitud = "PC"
    m_MockSolicitud.codigoSolicitud = "PC-001"

    ' Act
    Dim rutaGenerada As String
    rutaGenerada = docService.GenerarDocumento(m_MockSolicitud)

    ' Assert
    modAssert.AreNotEqual "", rutaGenerada, "La ruta del documento no debería estar vacía."
    modAssert.IsTrue InStr(rutaGenerada, "PC-001") > 0, "El nombre del fichero debe contener el código de la solicitud."

    result.Passed = True
    Exit Function

TestFail:
    result.Passed = False
    result.Message = "Error: " & Err.Description
End Function

Private Function Test_GenerarDocumento_PlantillaNoExiste_Fail() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_GenerarDocumento_PlantillaNoExiste_Fail"
    On Error GoTo TestFail

    ' Arrange
    Call Setup
    Dim docService As IDocumentService
    Set docService = CreateTestableDocumentService()

    ' Configurar mocks para un escenario de fallo (la plantilla no existe)
    m_MockConfig.SetPlantillaPath "C:RutaInexistente"
    m_MockConfig.IsTestMode = False ' Forzar la comprobación de ficheros para esta prueba
    m_MockSolicitud.tipoSolicitud = "PC"

    ' Act
    Dim rutaGenerada As String
    rutaGenerada = docService.GenerarDocumento(m_MockSolicitud)

    ' Assert
    modAssert.AreEqual "", rutaGenerada, "La ruta del documento debería estar vacía si la plantilla no existe."

    result.Passed = True
    Exit Function
TestFail:
    result.Passed = False
    result.Message = "Error: " & Err.Description
End Function

' --- FUNCIÓN AUXILIAR PARA CREAR EL SERVICIO BAJO PRUEBA ---
Private Function CreateTestableDocumentService() As IDocumentService
    ' Crea una instancia real del servicio pero le inyecta nuestros mocks
    Dim serviceImpl As New CDocumentService
    serviceImpl.Initialize m_MockConfig, m_MockSolicitudRepository, m_MockOperationLogger
    Set CreateTestableDocumentService = serviceImpl
End Function
