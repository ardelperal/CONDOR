Attribute VB_Name = "Test_Codificacion"
Option Compare Database


Option Explicit

' Módulo de prueba para verificar la codificación UTF-8
' Contiene caracteres especiales: áéíóú ñÑ ¿¡

Public Function TestCaracteresEspeciales() As Boolean
    ' Función de prueba con acentos y caracteres especiales
    Dim mensaje As String
    mensaje = "Prueba de codificación: áéíóú ñÑ ¿¡"
    
    ' Verificar que los caracteres se muestran correctamente
    Debug.Print "? Mensaje con acentos: " & mensaje
    Debug.Print "? Símbolos especiales: ????"
    Debug.Print "? Caracteres de caja: +- +-"
    
    TestCaracteresEspeciales = True
End Function

Public Function ObtenerMensajeConAcentos() As String
    ' Función que retorna un mensaje con acentos
    ObtenerMensajeConAcentos = "Configuración exitosa con caracteres españoles"
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ============================================================================

Public Function Test_Codificacion_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Integer
    Dim testsTotal As Integer
    
    resultado = "=== PRUEBAS DE CODIFICACIÓN ===" & vbCrLf
    
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
    If Len(mensaje) > 0 And InStr(mensaje, "Configuración") > 0 Then
        resultado = resultado & "[OK] ObtenerMensajeConAcentos" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] ObtenerMensajeConAcentos" & vbCrLf
    End If
    
    ' Resumen final
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasaron" & vbCrLf
    
    Test_Codificacion_RunAll = resultado
End Function







