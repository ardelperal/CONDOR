Attribute VB_Name = "TEST_SIMPLE"
' Módulo de prueba simple sin compilación condicional
Option Compare Database
Option Explicit

' Función de prueba muy simple
Public Function TEST_SIMPLE() As String
    TEST_SIMPLE = "Prueba simple sin acentos"
End Function

' Función que prueba operaciones básicas
Public Function TestBasico() As String
    Dim resultado As String
    resultado = "=== PRUEBA BASICA ===" & vbCrLf
    resultado = resultado & "✓ Función ejecutada correctamente" & vbCrLf
    resultado = resultado & "✓ Concatenación de strings: OK" & vbCrLf
    resultado = resultado & "=== FIN PRUEBA ===" & vbCrLf
    TestBasico = resultado
End Function
