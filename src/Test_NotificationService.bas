Attribute VB_Name = "Test_NotificationService"
' ===================================================================
' Módulo: Test_NotificationService
' Descripción: Pruebas unitarias para el servicio de notificaciones por correo electrónico
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================

Option Explicit

' Declaraciones a nivel de módulo
Private notificationService As INotificationService
Private mockSolicitud As T_Solicitud

' ===================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ===================================================================

Public Function Test_NotificationService_RunAll() As Boolean
    On Error GoTo ErrorHandler
    
    Dim totalTests As Long
    Dim passedTests As Long
    
    Debug.Print "=== INICIANDO PRUEBAS: NotificationService ==="
    
    ' Ejecutar todos los tests
    If Test_EnviarNotificacion_ConstruyeEmailCorrectamente() Then passedTests = passedTests + 1
    totalTests = totalTests + 1
    
    If Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect() Then passedTests = passedTests + 1
    totalTests = totalTests + 1
    
    If Test_EnviarNotificacion_SolicitudNula_RetornaFalso() Then passedTests = passedTests + 1
    totalTests = totalTests + 1
    
    ' Mostrar resumen
    Debug.Print "=== RESUMEN NotificationService: " & passedTests & "/" & totalTests & " pruebas pasaron ==="
    
    Test_NotificationService_RunAll = (passedTests = totalTests)
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en Test_NotificationService_RunAll: " & Err.Description
    Test_NotificationService_RunAll = False
End Function

' ===================================================================
' TESTS ESPECÍFICOS
' ===================================================================

Public Function Test_EnviarNotificacion_ConstruyeEmailCorrectamente() As Boolean
    On Error GoTo ErrorHandler
    
    ' Arrange
    Set notificationService = New CNotificationService
    Set mockSolicitud = New T_Solicitud
    
    With mockSolicitud
        .NumeroExpediente = "EXP-2024-001"
        .TipoSolicitud = "PC"
        .Descripcion = "Solicitud de cambio de precio"
        .Estado = "En Revisión"
        .UsuarioCreador = "test.user@empresa.com"
        .FechaCreacion = #1/15/2024#
    End With
    
    Dim tipoEvento As String
    tipoEvento = "CambioEstado"
    
    ' Act
    Dim resultado As Boolean
    Dim destinatario As String
    Dim asunto As String
    Dim cuerpo As String
    
    resultado = notificationService.EnviarNotificacion(mockSolicitud, tipoEvento, destinatario, asunto, cuerpo)
    
    ' Assert
    Dim testPassed As Boolean
    testPassed = True
    
    ' Verificar que se construyó el destinatario correctamente
    If destinatario <> "test.user@empresa.com" Then
        Debug.Print "✗ FAIL: Destinatario incorrecto. Esperado: test.user@empresa.com, Obtenido: " & destinatario
        testPassed = False
    End If
    
    ' Verificar que el asunto contiene información relevante
    If InStr(asunto, "EXP-2024-001") = 0 Or InStr(asunto, "En Revisión") = 0 Then
        Debug.Print "✗ FAIL: Asunto no contiene información esperada. Obtenido: " & asunto
        testPassed = False
    End If
    
    ' Verificar que el cuerpo contiene información de la solicitud
    If InStr(cuerpo, "EXP-2024-001") = 0 Or InStr(cuerpo, "Solicitud de cambio de precio") = 0 Then
        Debug.Print "✗ FAIL: Cuerpo no contiene información esperada. Obtenido: " & cuerpo
        testPassed = False
    End If
    
    If testPassed Then
        Debug.Print "✓ PASS: Test_EnviarNotificacion_ConstruyeEmailCorrectamente"
        Test_EnviarNotificacion_ConstruyeEmailCorrectamente = True
    Else
        Test_EnviarNotificacion_ConstruyeEmailCorrectamente = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "✗ ERROR: Test_EnviarNotificacion_ConstruyeEmailCorrectamente - " & Err.Description
    Test_EnviarNotificacion_ConstruyeEmailCorrectamente = False
End Function

Public Function Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect() As Boolean
    On Error GoTo ErrorHandler
    
    ' Arrange
    Set notificationService = New CNotificationService
    Set mockSolicitud = New T_Solicitud
    
    With mockSolicitud
        .NumeroExpediente = "EXP-2024-002"
        .TipoSolicitud = "RAC"
        .Estado = "Aprobada"
        .UsuarioCreador = "supervisor@empresa.com"
    End With
    
    ' Act
    Dim resultado As Boolean
    Dim destinatario As String
    Dim asunto As String
    Dim cuerpo As String
    
    resultado = notificationService.EnviarNotificacion(mockSolicitud, "CambioEstado", destinatario, asunto, cuerpo)
    
    ' Assert
    If InStr(asunto, "CONDOR") > 0 And InStr(asunto, "EXP-2024-002") > 0 And InStr(asunto, "Aprobada") > 0 Then
        Debug.Print "✓ PASS: Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect - Asunto: " & asunto
        Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect = True
    Else
        Debug.Print "✗ FAIL: Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect - Asunto: " & asunto
        Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "✗ ERROR: Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect - " & Err.Description
    Test_EnviarNotificacion_CambioEstado_GeneraAsuntoCorrect = False
End Function

Public Function Test_EnviarNotificacion_SolicitudNula_RetornaFalso() As Boolean
    On Error GoTo ErrorHandler
    
    ' Arrange
    Set notificationService = New CNotificationService
    Set mockSolicitud = Nothing  ' Solicitud nula
    
    ' Act
    Dim resultado As Boolean
    Dim destinatario As String
    Dim asunto As String
    Dim cuerpo As String
    
    resultado = notificationService.EnviarNotificacion(mockSolicitud, "CambioEstado", destinatario, asunto, cuerpo)
    
    ' Assert
    If resultado = False Then
        Debug.Print "✓ PASS: Test_EnviarNotificacion_SolicitudNula_RetornaFalso"
        Test_EnviarNotificacion_SolicitudNula_RetornaFalso = True
    Else
        Debug.Print "✗ FAIL: Test_EnviarNotificacion_SolicitudNula_RetornaFalso - Debería retornar False"
        Test_EnviarNotificacion_SolicitudNula_RetornaFalso = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "✗ ERROR: Test_EnviarNotificacion_SolicitudNula_RetornaFalso - " & Err.Description
    Test_EnviarNotificacion_SolicitudNula_RetornaFalso = False
End Function