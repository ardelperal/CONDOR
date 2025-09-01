Attribute VB_Name = "TestSolicitudService"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CSolicitudService
' Arquitectura: Pruebas Aisladas con Mocks
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function TestSolicitudServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TestSolicitudService - Pruebas Unitarias CSolicitudService")
    
    Call suiteResult.AddTestResult(TestCreateSolicitudSuccess())
    Call suiteResult.AddTestResult(TestCreateSolicitudFailsWithEmptyExpediente())
    Call suiteResult.AddTestResult(TestSaveSolicitudSuccess())
    
    Set TestSolicitudServiceRunAll = suiteResult
End Function



' ============================================================================
' PRUEBAS
' ============================================================================

Private Function TestCreateSolicitudSuccess() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("CreateSolicitud debe crear una solicitud con valores por defecto correctos")
    
    ' Variables locales
    Dim service As ISolicitudService
    Dim mockRepo As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim serviceImpl As CSolicitudService
    Dim expediente As EExpediente
    Dim result As ESolicitud
    Dim savedSolicitud As ESolicitud
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockSolicitudRepository
    mockRepo.Reset
    Set mockLogger = New CMockOperationLogger
    mockLogger.Reset
    Set mockErrorHandler = New CMockErrorHandlerService
    mockErrorHandler.Reset
    
    Dim serviceImpl As ISolicitudService
    Set serviceImpl = New CMockSolicitudService
    Call serviceImpl.Initialize(mockRepo, mockLogger, mockErrorHandler)
    Set service = serviceImpl
    
    Set expediente = New EExpediente
    expediente.idExpediente = "EXP001"
    
    Call mockRepo.ConfigureSaveSolicitud(123) ' Simular que el guardado devuelve un nuevo ID
    
    ' Act
    Set result = service.CreateSolicitud(expediente)
    
    ' Assert
    AssertNotNull result, "La solicitud devuelta no debe ser nula"
    AssertTrue mockRepo.SaveSolicitudCalled, "Se debe llamar al método SaveSolicitud del repositorio"
    
    Set savedSolicitud = mockRepo.LastSavedSolicitud
    
    AssertNotNull savedSolicitud, "El objeto solicitud debe haber sido pasado al repositorio"
    AssertEquals 1, savedSolicitud.idEstadoInterno, "El estado inicial debe ser 1 (Borrador)"
    AssertEquals expediente.idExpediente, savedSolicitud.idExpediente, "El idExpediente no es correcto"
    AssertEquals "PC", savedSolicitud.tipoSolicitud, "El tipoSolicitud no es correcto"
    AssertTrue InStr(savedSolicitud.codigoSolicitud, "PC-" & expediente.idExpediente) > 0, "El código de solicitud no tiene el formato esperado"
    AssertEquals Environ("USERNAME"), savedSolicitud.usuarioCreacion, "El usuario de creación no es el esperado"
    
    AssertEquals 123, result.idSolicitud, "El ID devuelto por el repo debe asignarse a la solicitud resultante"

    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    If Not mockRepo Is Nothing Then mockRepo.Reset
    If Not mockLogger Is Nothing Then mockLogger.Reset
    If Not mockErrorHandler Is Nothing Then mockErrorHandler.Reset
    Set service = Nothing
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set expediente = Nothing
    Set result = Nothing
    Set savedSolicitud = Nothing
    Set TestCreateSolicitudSuccess = testResult
End Function

Private Function TestCreateSolicitudFailsWithEmptyExpediente() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("CreateSolicitud debe fallar si idExpediente está vacío")
    
    ' Variables locales
    Dim service As ISolicitudService
    Dim mockRepo As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim serviceImpl As CSolicitudService
    Dim expediente As EExpediente
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockSolicitudRepository
    mockRepo.Reset
    Set mockLogger = New CMockOperationLogger
    mockLogger.Reset
    Set mockErrorHandler = New CMockErrorHandlerService
    mockErrorHandler.Reset

    Set serviceImpl = New CMockSolicitudService
    Call serviceImpl.Initialize(mockRepo, mockLogger, mockErrorHandler)
    Set service = serviceImpl
    
    Set expediente = New EExpediente
    expediente.idExpediente = " "
    
    ' Act & Assert
    On Error Resume Next
    Call service.CreateSolicitud(expediente)
    AssertEquals 5, Err.Number, "Debe lanzar un error si idExpediente está vacío"
    On Error GoTo TestFail
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("La prueba no debería haber llegado al manejador de errores principal")
Cleanup:
    If Not mockRepo Is Nothing Then mockRepo.Reset
    If Not mockLogger Is Nothing Then mockLogger.Reset
    If Not mockErrorHandler Is Nothing Then mockErrorHandler.Reset
    Set service = Nothing
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set expediente = Nothing
    Set TestCreateSolicitudFailsWithEmptyExpediente = testResult
End Function

Private Function TestSaveSolicitudSuccess() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("SaveSolicitud debe establecer los campos de modificación")
    
    ' Variables locales
    Dim service As ISolicitudService
    Dim mockRepo As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim serviceImpl As CSolicitudService
    Dim solicitud As ESolicitud
    Dim savedSolicitud As ESolicitud
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockSolicitudRepository
    mockRepo.Reset
    Set mockLogger = New CMockOperationLogger
    mockLogger.Reset
    Set mockErrorHandler = New CMockErrorHandlerService
    mockErrorHandler.Reset

    Set serviceImpl = New CMockSolicitudService
    Call serviceImpl.Initialize(mockRepo, mockLogger, mockErrorHandler)
    Set service = serviceImpl
    
    Set solicitud = New ESolicitud
    solicitud.idSolicitud = 456
    
    ' Act
    Call service.SaveSolicitud(solicitud)
    
    ' Assert
    AssertTrue mockRepo.SaveSolicitudCalled, "Se debe llamar al método SaveSolicitud del repositorio"
    
    Set savedSolicitud = mockRepo.LastSavedSolicitud
    
    AssertNotNull savedSolicitud, "El objeto solicitud debe haber sido pasado al repositorio"
    AssertEquals Environ("USERNAME"), savedSolicitud.usuarioModificacion, "El usuario de modificación no es el esperado"
    AssertTrue IsDate(savedSolicitud.fechaModificacion), "La fecha de modificación debe ser una fecha válida"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    If Not mockRepo Is Nothing Then mockRepo.Reset
    If Not mockLogger Is Nothing Then mockLogger.Reset
    If Not mockErrorHandler Is Nothing Then mockErrorHandler.Reset
    Set service = Nothing
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set solicitud = Nothing
    Set savedSolicitud = Nothing
    Set TestSaveSolicitudSuccess = testResult
End Function