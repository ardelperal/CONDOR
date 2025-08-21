Option Compare Database
Option Explicit
' Test_NotificationService.bas
' Módulo de pruebas para el servicio de notificaciones
' Sigue el patrón TDD y las lecciones aprendidas del proyecto CONDOR
' Basado en las Especificaciones de Integración - Sección 2


' Variables de módulo para el mock
Private mockNotificationService As INotificationService

' Función principal para ejecutar todas las pruebas del módulo
Public Function Test_NotificationService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.SuiteName = "Test_NotificationService"
    
    Debug.Print "=== Iniciando Test_NotificationService ==="
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTest "Test_EnviarNotificacion_ParametrosValidos", Test_EnviarNotificacion_ParametrosValidos_Result()
    suiteResult.AddTest "Test_EnviarNotificacion_ParametrosInvalidos", Test_EnviarNotificacion_ParametrosInvalidos_Result()
    suiteResult.AddTest "Test_EnviarNotificacion_ConAdjunto", Test_EnviarNotificacion_ConAdjunto_Result()
    suiteResult.AddTest "Test_EnviarNotificacion_SinAdjunto", Test_EnviarNotificacion_SinAdjunto_Result()
    suiteResult.AddTest "Test_EnviarNotificacion_SimularError", Test_EnviarNotificacion_SimularError_Result()
    
    Debug.Print "=== Test_NotificationService Completado ==="
    
    Set Test_NotificationService_RunAll = suiteResult
End Function

' Test: Verificar envío con parámetros válidos
Private Function Test_EnviarNotificacion_ParametrosValidos_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_ParametrosValidos"
    
    On Error GoTo TestError
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.ValorRetorno = True
    
    Dim destinatarios As String
    Dim asunto As String
    Dim cuerpoHTML As String
    Dim urlAdjunto As String
    
    destinatarios = "usuario1@empresa.com;usuario2@empresa.com"
    asunto = "Notificación de Prueba CONDOR"
    cuerpoHTML = "<html><body><h1>Prueba</h1><p>Contenido de prueba</p></body></html>"
    urlAdjunto = "C:\temp\documento.pdf"
    
    ' Act
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion(destinatarios, asunto, cuerpoHTML, urlAdjunto)
    
    ' Assert
    Assert.IsTrue resultado, "El método debe devolver True con parámetros válidos"
    Assert.AreEqual 1, mockNotificationService.NumeroLlamadas, "Debe registrar exactamente 1 llamada"
    Assert.IsTrue mockNotificationService.FueLlamadoCon(destinatarios, asunto, cuerpoHTML, urlAdjunto), "Los parámetros no coinciden con la llamada registrada"
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion con parámetros válidos"
    
    Set Test_EnviarNotificacion_ParametrosValidos_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_ParametrosValidos: " & Err.Description
    Set Test_EnviarNotificacion_ParametrosValidos_Result = result
End Function

' Test: Verificar comportamiento con parámetros inválidos
Private Function Test_EnviarNotificacion_ParametrosInvalidos_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_ParametrosInvalidos"
    
    On Error GoTo TestError
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.ValorRetorno = False
    
    ' Act & Assert - Destinatarios vacío
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion("", "Asunto", "<html>Cuerpo</html>")
    
    Assert.AreEqual 1, mockNotificationService.NumeroLlamadas, "Debe registrar la llamada aunque los parámetros sean inválidos"
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion con parámetros inválidos"
    
    Set Test_EnviarNotificacion_ParametrosInvalidos_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_ParametrosInvalidos: " & Err.Description
    Set Test_EnviarNotificacion_ParametrosInvalidos_Result = result
End Function

' Test: Verificar envío con adjunto
Private Function Test_EnviarNotificacion_ConAdjunto_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_ConAdjunto"
    
    On Error GoTo TestError
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.ValorRetorno = True
    
    Dim urlAdjunto As String
    urlAdjunto = "C:\documentos\reporte.pdf"
    
    ' Act
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>", urlAdjunto)
    
    ' Assert
    Assert.IsTrue resultado, "El método debe devolver True con adjunto"
    Assert.AreEqual urlAdjunto, mockNotificationService.UltimaLlamada_URLAdjunto, "URL de adjunto no registrada correctamente"
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion con adjunto"
    
    Set Test_EnviarNotificacion_ConAdjunto_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_ConAdjunto: " & Err.Description
    Set Test_EnviarNotificacion_ConAdjunto_Result = result
End Function

' Test: Verificar envío sin adjunto
Private Function Test_EnviarNotificacion_SinAdjunto_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_SinAdjunto"
    
    On Error GoTo TestError
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.ValorRetorno = True
    
    ' Act
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Assert
    Assert.IsTrue resultado, "El método debe devolver True sin adjunto"
    Assert.AreEqual "", mockNotificationService.UltimaLlamada_URLAdjunto, "URL de adjunto debe estar vacía cuando no se proporciona"
    
    result.Passed = True
    result.Message = "Prueba exitosa: EnviarNotificacion sin adjunto"
    
    Set Test_EnviarNotificacion_SinAdjunto_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_SinAdjunto: " & Err.Description
    Set Test_EnviarNotificacion_SinAdjunto_Result = result
End Function

' Test: Verificar manejo de errores
Private Function Test_EnviarNotificacion_SimularError_Result() As CTestResult
    Dim result As New CTestResult
    result.TestName = "Test_EnviarNotificacion_SimularError"
    
    On Error GoTo TestError
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.SimularError = True
    
    ' Act & Assert
    On Error GoTo ErrorEsperado
    
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Si llegamos aquí, no se produjo el error esperado
    result.Passed = False
    result.Message = "ERROR: Se esperaba un error simulado pero no ocurrió"
    Set Test_EnviarNotificacion_SimularError_Result = result
    Exit Function
    
ErrorEsperado:
    If Err.Number = 9999 Then
        result.Passed = True
        result.Message = "Prueba exitosa: Error simulado manejado correctamente"
    Else
        result.Passed = False
        result.Message = "ERROR: Error inesperado: " & Err.Description
    End If
    
    On Error GoTo 0
    Set Test_EnviarNotificacion_SimularError_Result = result
    Exit Function
    
TestError:
    result.Passed = False
    result.Message = "Error en Test_EnviarNotificacion_SimularError: " & Err.Description
    Set Test_EnviarNotificacion_SimularError_Result = result
End Function






