Option Compare Database
Option Explicit
' ===================================================================
' Módulo: Test_ValidationService
' Descripción: Pruebas unitarias para el servicio de validaciones de negocio
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================


' Declaraciones a nivel de módulo
Private validationService As IValidationService
Private mockSolicitud As T_Solicitud

' ===================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ===================================================================

Public Function Test_ValidationService_RunAll() As Boolean
    On Error GoTo ErrorHandler
    
    Dim totalTests As Long
    Dim passedTests As Long
    
    Debug.Print "=== INICIANDO PRUEBAS: ValidationService ==="
    
    ' Ejecutar todos los tests
    If Test_ValidarSolicitud_ConDatosCompletos_RetornaExito() Then passedTests = passedTests + 1
    totalTests = totalTests + 1
    
    If Test_ValidarSolicitud_SinExpediente_RetornaFallo() Then passedTests = passedTests + 1
    totalTests = totalTests + 1
    
    ' Mostrar resumen
    Debug.Print "=== RESUMEN ValidationService: " & passedTests & "/" & totalTests & " pruebas pasaron ==="
    
    Test_ValidationService_RunAll = (passedTests = totalTests)
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en Test_ValidationService_RunAll: " & Err.Description
    Test_ValidationService_RunAll = False
End Function

' ===================================================================
' TESTS ESPECÍFICOS
' ===================================================================

Public Function Test_ValidarSolicitud_ConDatosCompletos_RetornaExito() As Boolean
    On Error GoTo ErrorHandler
    
    Dim localValidationService As IValidationService
    Dim localMockSolicitud As T_Solicitud
    
    ' Arrange
    Set localValidationService = New CValidationService
    Set localMockSolicitud = New T_Solicitud
    
    With localMockSolicitud
        .NumeroExpediente = "EXP-2024-001"
        .tipoSolicitud = "PC"
        .Descripcion = "Solicitud de cambio de precio válida"
        .justificacionCambio = "Incremento de materiales según índice ICCP"
        .importeOriginal = 1500000
        .importeNuevo = 1725000
        .Estado = "Borrador"
        .fechaCreacion = Now
        .UsuarioCreador = "test.user@empresa.com"
    End With
    
    ' Act
    Dim resultado As Boolean
    Dim MensajeError As String
    resultado = localValidationService.ValidarSolicitud(localMockSolicitud, MensajeError)
    
    ' Assert
    If resultado = True And MensajeError = "" Then
        Debug.Print "✓ PASS: Test_ValidarSolicitud_ConDatosCompletos_RetornaExito"
        Test_ValidarSolicitud_ConDatosCompletos_RetornaExito = True
    Else
        Debug.Print "✗ FAIL: Test_ValidarSolicitud_ConDatosCompletos_RetornaExito - Resultado: " & resultado & ", Error: " & MensajeError
        Test_ValidarSolicitud_ConDatosCompletos_RetornaExito = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "✗ ERROR: Test_ValidarSolicitud_ConDatosCompletos_RetornaExito - " & Err.Description
    Test_ValidarSolicitud_ConDatosCompletos_RetornaExito = False
End Function

Public Function Test_ValidarSolicitud_SinExpediente_RetornaFallo() As Boolean
    On Error GoTo ErrorHandler
    
    Dim localValidationService As IValidationService
    Dim localMockSolicitud As T_Solicitud
    
    ' Arrange
    Set localValidationService = New CValidationService
    Set localMockSolicitud = New T_Solicitud
    
    With localMockSolicitud
        .NumeroExpediente = ""  ' Campo vacío - debe fallar
        .tipoSolicitud = "PC"
        .Descripcion = "Solicitud sin expediente"
        .justificacionCambio = "Test de validación"
        .importeOriginal = 1500000
        .importeNuevo = 1725000
        .Estado = "Borrador"
        .fechaCreacion = Now
        .UsuarioCreador = "test.user@empresa.com"
    End With
    
    ' Act
    Dim resultado As Boolean
    Dim MensajeError As String
    resultado = localValidationService.ValidarSolicitud(localMockSolicitud, MensajeError)
    
    ' Assert
    If resultado = False And InStr(MensajeError, "expediente") > 0 Then
        Debug.Print "✓ PASS: Test_ValidarSolicitud_SinExpediente_RetornaFallo - Error: " & MensajeError
        Test_ValidarSolicitud_SinExpediente_RetornaFallo = True
    Else
        Debug.Print "✗ FAIL: Test_ValidarSolicitud_SinExpediente_RetornaFallo - Resultado: " & resultado & ", Error: " & MensajeError
        Test_ValidarSolicitud_SinExpediente_RetornaFallo = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "✗ ERROR: Test_ValidarSolicitud_SinExpediente_RetornaFallo - " & Err.Description
    Test_ValidarSolicitud_SinExpediente_RetornaFallo = False
End Function






