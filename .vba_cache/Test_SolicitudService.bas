Attribute VB_Name = "Test_SolicitudService"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CSolicitudService
' Arquitectura: Pruebas Aisladas con Mocks
' ============================================================================

Private m_service As ISolicitudService
Private m_mockRepo As CMockSolicitudRepository
Private m_mockLogger As CMockOperationLogger
Private m_mockErrorHandler As CMockErrorHandlerService

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function Test_SolicitudService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("Test_SolicitudService - Pruebas Unitarias CSolicitudService")
    
    Call suiteResult.AddTestResult(Test_CreateSolicitud_Success())
    Call suiteResult.AddTestResult(Test_CreateSolicitud_FailsWithEmptyExpediente())
    Call suiteResult.AddTestResult(Test_SaveSolicitud_Success())
    
    Set Test_SolicitudService_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Set m_mockRepo = New CMockSolicitudRepository
    Set m_mockLogger = New CMockOperationLogger
    Set m_mockErrorHandler = New CMockErrorHandlerService
    
    Dim serviceImpl As New CSolicitudService
    Call serviceImpl.Initialize(m_mockRepo, m_mockLogger, m_mockErrorHandler)
    Set m_service = serviceImpl
End Sub

Private Sub Teardown()
    m_mockRepo.Reset
    m_mockLogger.Reset
    m_mockErrorHandler.Reset
    Set m_service = Nothing
    Set m_mockRepo = Nothing
    Set m_mockLogger = Nothing
    Set m_mockErrorHandler = Nothing
End Sub

' ============================================================================
' PRUEBAS
' ============================================================================

Private Function Test_CreateSolicitud_Success() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("CreateSolicitud debe crear una solicitud con valores por defecto correctos")
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim idExpediente As String
    idExpediente = "EXP001"
    Dim tipo As String
    tipo = "PC"
    
    Call m_mockRepo.SetSaveSolicitudReturnValue(123) ' Simular que el guardado devuelve un nuevo ID
    
    ' Act
    Dim result As T_Solicitud
    Set result = m_service.CreateSolicitud(idExpediente, tipo)
    
    ' Assert
    AssertNotNull result, "La solicitud devuelta no debe ser nula"
    AssertTrue m_mockRepo.SaveSolicitudCalled, "Se debe llamar al método SaveSolicitud del repositorio"
    
    Dim savedSolicitud As T_Solicitud
    Set savedSolicitud = m_mockRepo.LastSavedSolicitud
    
    AssertNotNull savedSolicitud, "El objeto solicitud debe haber sido pasado al repositorio"
    AssertEquals 1, savedSolicitud.idEstadoInterno, "El estado inicial debe ser 1 (Borrador)"
    AssertEquals idExpediente, savedSolicitud.idExpediente, "El idExpediente no es correcto"
    AssertEquals tipo, savedSolicitud.tipoSolicitud, "El tipoSolicitud no es correcto"
    AssertTrue InStr(savedSolicitud.codigoSolicitud, tipo & "-" & idExpediente) > 0, "El código de solicitud no tiene el formato esperado"
    AssertEquals Environ("USERNAME"), savedSolicitud.usuarioCreacion, "El usuario de creación no es el esperado"
    
    AssertEquals 123, result.idSolicitud, "El ID devuelto por el repo debe asignarse a la solicitud resultante"

    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    Call Teardown
    Set Test_CreateSolicitud_Success = testResult
End Function

Private Function Test_CreateSolicitud_FailsWithEmptyExpediente() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("CreateSolicitud debe fallar si idExpediente está vacío")
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Act & Assert
    On Error Resume Next
    Call m_service.CreateSolicitud(" ", "PC")
    AssertEquals 5, Err.Number, "Debe lanzar un error si idExpediente está vacío"
    On Error GoTo ErrorHandler
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    Call testResult.Fail("La prueba no debería haber llegado al manejador de errores principal")
Cleanup:
    Call Teardown
    Set Test_CreateSolicitud_FailsWithEmptyExpediente = testResult
End Function

Private Function Test_SaveSolicitud_Success() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("SaveSolicitud debe establecer los campos de modificación")
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim solicitud As New T_Solicitud
    solicitud.idSolicitud = 456
    
    ' Act
    Call m_service.SaveSolicitud(solicitud)
    
    ' Assert
    AssertTrue m_mockRepo.SaveSolicitudCalled, "Se debe llamar al método SaveSolicitud del repositorio"
    
    Dim savedSolicitud As T_Solicitud
    Set savedSolicitud = m_mockRepo.LastSavedSolicitud
    
    AssertNotNull savedSolicitud, "El objeto solicitud debe haber sido pasado al repositorio"
    AssertEquals Environ("USERNAME"), savedSolicitud.usuarioModificacion, "El usuario de modificación no es el esperado"
    AssertTrue IsDate(savedSolicitud.fechaModificacion), "La fecha de modificación debe ser una fecha válida"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_SaveSolicitud_Success = testResult
End Function
