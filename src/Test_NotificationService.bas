Attribute VB_Name = "Test_NotificationService"
Option Compare Database


Option Explicit

' Test_NotificationService.bas
' Módulo de pruebas para el servicio de notificaciones
' Sigue el patrón TDD y las lecciones aprendidas del proyecto CONDOR
' Basado en las Especificaciones de Integración - Sección 2


' Variables de módulo para el mock
Private mockNotificationService As INotificationService

' Función principal para ejecutar todas las pruebas del módulo
Public Sub Test_NotificationService_RunAll()
    Debug.Print "=== Iniciando Test_NotificationService ==="
    
    ' Ejecutar todas las pruebas
    Test_EnviarNotificacion_ParametrosValidos
    Test_EnviarNotificacion_ParametrosInvalidos
    Test_EnviarNotificacion_ConAdjunto
    Test_EnviarNotificacion_SinAdjunto
    Test_EnviarNotificacion_SimularError
    
    Debug.Print "=== Test_NotificationService Completado ==="
End Sub

' Test: Verificar envío con parámetros válidos
Public Sub Test_EnviarNotificacion_ParametrosValidos()
    Debug.Print "Ejecutando: Test_EnviarNotificacion_ParametrosValidos"
    
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
    If Not resultado Then
        Debug.Print "ERROR: El método debe devolver True con parámetros válidos"
        Exit Sub
    End If
    
    If mockNotificationService.NumeroLlamadas <> 1 Then
        Debug.Print "ERROR: Debe registrar exactamente 1 llamada"
        Exit Sub
    End If
    
    If Not mockNotificationService.FueLlamadoCon(destinatarios, asunto, cuerpoHTML, urlAdjunto) Then
        Debug.Print "ERROR: Los parámetros no coinciden con la llamada registrada"
        Exit Sub
    End If
    
    Debug.Print "ÉXITO: Test_EnviarNotificacion_ParametrosValidos"
End Sub

' Test: Verificar comportamiento con parámetros inválidos
Public Sub Test_EnviarNotificacion_ParametrosInvalidos()
    Debug.Print "Ejecutando: Test_EnviarNotificacion_ParametrosInvalidos"
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.ValorRetorno = False
    
    ' Act & Assert - Destinatarios vacío
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion("", "Asunto", "<html>Cuerpo</html>")
    
    If mockNotificationService.NumeroLlamadas <> 1 Then
        Debug.Print "ERROR: Debe registrar la llamada aunque los parámetros sean inválidos"
        Exit Sub
    End If
    
    Debug.Print "ÉXITO: Test_EnviarNotificacion_ParametrosInvalidos"
End Sub

' Test: Verificar envío con adjunto
Public Sub Test_EnviarNotificacion_ConAdjunto()
    Debug.Print "Ejecutando: Test_EnviarNotificacion_ConAdjunto"
    
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
    If Not resultado Then
        Debug.Print "ERROR: El método debe devolver True con adjunto"
        Exit Sub
    End If
    
    If mockNotificationService.UltimaLlamada_URLAdjunto <> urlAdjunto Then
        Debug.Print "ERROR: URL de adjunto no registrada correctamente"
        Exit Sub
    End If
    
    Debug.Print "ÉXITO: Test_EnviarNotificacion_ConAdjunto"
End Sub

' Test: Verificar envío sin adjunto
Public Sub Test_EnviarNotificacion_SinAdjunto()
    Debug.Print "Ejecutando: Test_EnviarNotificacion_SinAdjunto"
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.ValorRetorno = True
    
    ' Act
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Assert
    If Not resultado Then
        Debug.Print "ERROR: El método debe devolver True sin adjunto"
        Exit Sub
    End If
    
    If mockNotificationService.UltimaLlamada_URLAdjunto <> "" Then
        Debug.Print "ERROR: URL de adjunto debe estar vacía cuando no se proporciona"
        Exit Sub
    End If
    
    Debug.Print "ÉXITO: Test_EnviarNotificacion_SinAdjunto"
End Sub

' Test: Verificar manejo de errores
Public Sub Test_EnviarNotificacion_SimularError()
    Debug.Print "Ejecutando: Test_EnviarNotificacion_SimularError"
    
    ' Arrange
    Set mockNotificationService = New CMockNotificationService
    mockNotificationService.LimpiarHistorial
    mockNotificationService.SimularError = True
    
    ' Act & Assert
    On Error GoTo ErrorEsperado
    
    Dim resultado As Boolean
    resultado = mockNotificationService.EnviarNotificacion("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Si llegamos aquí, no se produjo el error esperado
    Debug.Print "ERROR: Se esperaba un error simulado"
    Exit Sub
    
ErrorEsperado:
    If Err.Number = 9999 Then
        Debug.Print "ÉXITO: Test_EnviarNotificacion_SimularError"
    Else
        Debug.Print "ERROR: Error inesperado: " & Err.Description
    End If
    
    On Error GoTo 0
End Sub





