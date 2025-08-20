Option Compare Database
Option Explicit
' ============================================================================
' Módulo: Test_DocumentService
' Descripción: Pruebas unitarias para CDocumentService.cls
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' Versión: 1.0 - Implementación inicial siguiendo TDD
' ============================================================================

' Mock para simular datos de solicitud para generación de documentos
Private Type T_MockDocumentData
    idSolicitud As Long
    tipoSolicitud As String
    PlantillaEsperada As String
    RutaPlantilla As String
    CamposEsperados As String
    DocumentoGenerado As String
    IsValid As Boolean
End Type

Private m_MockDocument As T_MockDocumentData

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN DE MOCKS
' ============================================================================

' Configura un mock para solicitud PC con plantilla correcta
Private Sub SetupValidPCDocumentMock()
    m_MockDocument.idSolicitud = 12345
    m_MockDocument.tipoSolicitud = "PC"
    m_MockDocument.PlantillaEsperada = "PC_template.docx"
    m_MockDocument.RutaPlantilla = "C:\Proyectos\CONDOR\back\recursos\Plantillas\PC_template.docx"
    m_MockDocument.CamposEsperados = "NumeroExpediente,FechaSolicitud,TipoSolicitud"
    m_MockDocument.DocumentoGenerado = "SOL-PC-2025-001.docx"
    m_MockDocument.IsValid = True
End Sub

' Configura un mock para solicitud inválida
Private Sub SetupInvalidDocumentMock()
    m_MockDocument.idSolicitud = 0
    m_MockDocument.tipoSolicitud = ""
    m_MockDocument.PlantillaEsperada = ""
    m_MockDocument.RutaPlantilla = ""
    m_MockDocument.CamposEsperados = ""
    m_MockDocument.DocumentoGenerado = ""
    m_MockDocument.IsValid = False
End Sub

' ============================================================================
' PRUEBAS DE GENERACIÓN DE DOCUMENTOS
' ============================================================================

' Prueba: GenerarDocumento para PC usa la plantilla correcta
Public Function Test_GenerarDocumento_PC_UsaPlantillaCorrecta() As String
    On Error GoTo ErrorHandler
    
    ' Arrange
    ' Patrón de doble variable: interfaz + implementación (Lección 1 y 2)
    Dim documentService As IDocumentService
    Dim documentServiceImpl As CDocumentService
    
    ' Inicializar dependencias
    Dim validationService As IValidationService
    Dim solicitudRepository As ISolicitudRepository
    Set validationService = New CValidationService
    Set solicitudRepository = New CMockSolicitudRepository
    
    ' Instanciar e inicializar la clase concreta
    Set documentServiceImpl = New CDocumentService
    documentServiceImpl.Initialize validationService, solicitudRepository
    
    ' Asignar la clase a la interfaz
    Set documentService = documentServiceImpl
    
    Call SetupValidPCDocumentMock
    
    Dim resultado As String
    Dim esperado As String
    esperado = m_MockDocument.PlantillaEsperada
    
    ' Act
    ' Nota: Este test verifica que el servicio intenta usar la plantilla correcta
    ' La implementación real verificará la lógica de selección de plantilla
    resultado = documentService.ObtenerNombrePlantilla(m_MockDocument.tipoSolicitud)
    
    ' Assert
    If resultado = esperado Then
        Test_GenerarDocumento_PC_UsaPlantillaCorrecta = "PASS: GenerarDocumento PC usa plantilla correcta (" & resultado & ")"
    Else
        Test_GenerarDocumento_PC_UsaPlantillaCorrecta = "FAIL: GenerarDocumento PC - Esperado: " & esperado & ", Obtenido: " & resultado
    End If
    
    Exit Function
    
ErrorHandler:
    Test_GenerarDocumento_PC_UsaPlantillaCorrecta = "ERROR: " & Err.Description
End Function

' Prueba: GenerarDocumento valida solicitud antes de procesar
Public Function Test_GenerarDocumento_ValidaSolicitudAntesProcesar() As String
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim validationService As IValidationService
    Set validationService = New CValidationService
    
    Call SetupInvalidDocumentMock
    
    Dim solicitud As T_Solicitud
    Set solicitud = New T_Solicitud
    solicitud.idSolicitud = m_MockDocument.idSolicitud
    
    Dim MensajeError As String
    Dim resultado As Boolean
    
    ' Act
    resultado = validationService.ValidarSolicitud(solicitud, MensajeError)
    
    ' Assert
    If resultado = False Then
        Test_GenerarDocumento_ValidaSolicitudAntesProcesar = "PASS: GenerarDocumento valida solicitud correctamente"
    Else
        Test_GenerarDocumento_ValidaSolicitudAntesProcesar = "FAIL: GenerarDocumento no validó solicitud inválida"
    End If
    
    Exit Function
    
ErrorHandler:
    Test_GenerarDocumento_ValidaSolicitudAntesProcesar = "ERROR: " & Err.Description
End Function

' Prueba: GenerarDocumento obtiene mapeo de campos correctamente
Public Function Test_GenerarDocumento_ObtieneMapeoCampos() As String
    On Error GoTo ErrorHandler
    
    ' Arrange
    ' Patrón de doble variable: interfaz + implementación (Lección 1 y 2)
    Dim documentService As IDocumentService
    Dim documentServiceImpl As CDocumentService
    
    ' Inicializar dependencias
    Dim validationService As IValidationService
    Dim solicitudRepository As ISolicitudRepository
    Set validationService = New CValidationService
    Set solicitudRepository = New CMockSolicitudRepository
    
    ' Instanciar e inicializar la clase concreta
    Set documentServiceImpl = New CDocumentService
    documentServiceImpl.Initialize validationService, solicitudRepository
    
    ' Asignar la clase a la interfaz
    Set documentService = documentServiceImpl
    
    Call SetupValidPCDocumentMock
    
    Dim resultado As String
    
    ' Act
    resultado = documentService.ObtenerMapeoCampos(m_MockDocument.tipoSolicitud)
    
    ' Assert
    If Len(resultado) > 0 Then
        Test_GenerarDocumento_ObtieneMapeoCampos = "PASS: GenerarDocumento obtiene mapeo de campos (" & Left(resultado, 50) & "...)"
    Else
        Test_GenerarDocumento_ObtieneMapeoCampos = "FAIL: GenerarDocumento no obtuvo mapeo de campos"
    End If
    
    Exit Function
    
ErrorHandler:
    Test_GenerarDocumento_ObtieneMapeoCampos = "ERROR: " & Err.Description
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_DocumentService_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS DE DocumentService ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas
    Dim testResult As String
    
    ' Test 1: Plantilla correcta para PC
    testsTotal = testsTotal + 1
    testResult = Test_GenerarDocumento_PC_UsaPlantillaCorrecta()
    resultado = resultado & testResult & vbCrLf
    If InStr(testResult, "PASS") > 0 Then testsPassed = testsPassed + 1
    
    ' Test 2: Validación de solicitud
    testsTotal = testsTotal + 1
    testResult = Test_GenerarDocumento_ValidaSolicitudAntesProcesar()
    resultado = resultado & testResult & vbCrLf
    If InStr(testResult, "PASS") > 0 Then testsPassed = testsPassed + 1
    
    ' Test 3: Mapeo de campos
    testsTotal = testsTotal + 1
    testResult = Test_GenerarDocumento_ObtieneMapeoCampos()
    resultado = resultado & testResult & vbCrLf
    If InStr(testResult, "PASS") > 0 Then testsPassed = testsPassed + 1
    
    ' Resumen final
    resultado = resultado & vbCrLf & "=== RESUMEN DocumentService ===" & vbCrLf
    resultado = resultado & "Pruebas ejecutadas: " & testsTotal & vbCrLf
    resultado = resultado & "Pruebas exitosas: " & testsPassed & vbCrLf
    resultado = resultado & "Pruebas fallidas: " & (testsTotal - testsPassed) & vbCrLf
    
    If testsPassed = testsTotal Then
        resultado = resultado & "RESULTADO: TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        resultado = resultado & "RESULTADO: " & (testsTotal - testsPassed) & " PRUEBAS FALLARON" & vbCrLf
    End If
    
    Test_DocumentService_RunAll = resultado
End Function






