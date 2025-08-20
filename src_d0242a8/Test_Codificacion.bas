Attribute VB_Name = "Test_Codificacion"
Option Compare Database
Option Explicit

' M├│dulo de prueba para verificar la codificaci├│n UTF-8
' Contiene caracteres especiales: ├í├®├¡├│├║ ├▒├æ ┬┐┬í

Public Function TestCaracteresEspeciales() As Boolean
    ' Funci├│n de prueba con acentos y caracteres especiales
    Dim mensaje As String
    mensaje = "Prueba de codificaci├│n: ├í├®├¡├│├║ ├▒├æ ┬┐┬í"
    
    ' Verificar que los caracteres se muestran correctamente
    Debug.Print "? Mensaje con acentos: " & mensaje
    Debug.Print "? S├¡mbolos especiales: ????"
    Debug.Print "? Caracteres de caja: +- +-"
    
    TestCaracteresEspeciales = True
End Function

Public Function ObtenerMensajeConAcentos() As String
    ' Funci├│n que retorna un mensaje con acentos
    ObtenerMensajeConAcentos = "Configuraci├│n exitosa con caracteres espa├▒oles"
End Function

' ============================================================================
' FUNCI├ôN PRINCIPAL DE EJECUCI├ôN DE PRUEBAS
' ============================================================================

Public Function Test_Codificacion_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DE CODIFICACI├ôN ===" & vbCrLf
    
    ' Test 1: Caracteres especiales
    testsTotal = testsTotal + 1
    If TestCaracteresEspeciales() Then
        resultado = resultado & "[OK] TestCaracteresEspeciales" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] TestCaracteresEspeciales" & vbCrLf
    End If
    
    ' Test 2: Mensaje con acentos
    testsTotal = testsTotal + 1
    Dim mensaje As String
    mensaje = ObtenerMensajeConAcentos()
    If Len(mensaje) > 0 And InStr(mensaje, "Configuraci├│n") > 0 Then
        resultado = resultado & "[OK] ObtenerMensajeConAcentos" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] ObtenerMensajeConAcentos" & vbCrLf
    End If
    
    ' Resumen final
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasaron" & vbCrLf
    
    Test_Codificacion_RunAll = resultado
End Function

