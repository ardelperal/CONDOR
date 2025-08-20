Attribute VB_Name = "Test_ValidationService"
' ===================================================================
' M├│dulo: Test_ValidationService
' Descripci├│n: Pruebas unitarias para el servicio de validaciones de negocio
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================

Option Explicit

' Declaraciones a nivel de m├│dulo
Private validationService As IValidationService
Private mockSolicitud As T_Solicitud

' ===================================================================
' FUNCI├ôN PRINCIPAL DE EJECUCI├ôN DE PRUEBAS
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
' TESTS ESPEC├ìFICOS
' ===================================================================

Public Function Test_ValidarSolicitud_ConDatosCompletos_RetornaExito() As Boolean
    On Error GoTo ErrorHandler
    
    ' Arrange
    Set validationService = New CValidationService
    Set mockSolicitud = New T_Solicitud
    
    With mockSolicitud
        .NumeroExpediente = "EXP-2024-001"
        .TipoSolicitud = "PC"
        .Descripcion = "Solicitud de cambio de precio v├ílida"
        .Justificacion = "Incremento de materiales seg├║n ├¡ndice ICCP"
        .ImporteOriginal = 1500000
        .ImporteNuevo = 1725000
        .Estado = "Borrador"
        .FechaCreacion = Now
        .UsuarioCreador = "test.user@empresa.com"
    End With
    
    ' Act
    Dim resultado As Boolean
    Dim mensajeError As String
    resultado = validationService.ValidarSolicitud(mockSolicitud, mensajeError)
    
    ' Assert
    If resultado = True And mensajeError = "" Then
        Debug.Print "Ô£ô PASS: Test_ValidarSolicitud_ConDatosCompletos_RetornaExito"
        Test_ValidarSolicitud_ConDatosCompletos_RetornaExito = True
    Else
        Debug.Print "Ô£ù FAIL: Test_ValidarSolicitud_ConDatosCompletos_RetornaExito - Resultado: " & resultado & ", Error: " & mensajeError
        Test_ValidarSolicitud_ConDatosCompletos_RetornaExito = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Ô£ù ERROR: Test_ValidarSolicitud_ConDatosCompletos_RetornaExito - " & Err.Description
    Test_ValidarSolicitud_ConDatosCompletos_RetornaExito = False
End Function

Public Function Test_ValidarSolicitud_SinExpediente_RetornaFallo() As Boolean
    On Error GoTo ErrorHandler
    
    ' Arrange
    Set validationService = New CValidationService
    Set mockSolicitud = New T_Solicitud
    
    With mockSolicitud
        .NumeroExpediente = ""  ' Campo vac├¡o - debe fallar
        .TipoSolicitud = "PC"
        .Descripcion = "Solicitud sin expediente"
        .Justificacion = "Test de validaci├│n"
        .ImporteOriginal = 1500000
        .ImporteNuevo = 1725000
        .Estado = "Borrador"
        .FechaCreacion = Now
        .UsuarioCreador = "test.user@empresa.com"
    End With
    
    ' Act
    Dim resultado As Boolean
    Dim mensajeError As String
    resultado = validationService.ValidarSolicitud(mockSolicitud, mensajeError)
    
    ' Assert
    If resultado = False And InStr(mensajeError, "expediente") > 0 Then
        Debug.Print "Ô£ô PASS: Test_ValidarSolicitud_SinExpediente_RetornaFallo - Error: " & mensajeError
        Test_ValidarSolicitud_SinExpediente_RetornaFallo = True
    Else
        Debug.Print "Ô£ù FAIL: Test_ValidarSolicitud_SinExpediente_RetornaFallo - Resultado: " & resultado & ", Error: " & mensajeError
        Test_ValidarSolicitud_SinExpediente_RetornaFallo = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Ô£ù ERROR: Test_ValidarSolicitud_SinExpediente_RetornaFallo - " & Err.Description
    Test_ValidarSolicitud_SinExpediente_RetornaFallo = False
End Function
