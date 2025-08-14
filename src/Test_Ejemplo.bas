Attribute VB_Name = "Test_Ejemplo"
Option Compare Database
Option Explicit

' Modulo de pruebas de ejemplo para demostrar el funcionamiento del motor de pruebas
' Todas las pruebas deben empezar con "Test_" para ser detectadas automaticamente

Public Sub Test_SumaBasica()
    ' Prueba que siempre pasa - suma basica
    Dim resultado As Integer
    resultado = 2 + 2
    
    ' Esta prueba pasa silenciosamente - no genera errores
End Sub

Public Sub Test_ConcatenacionTexto()
    ' Prueba que siempre pasa - concatenacion de texto
    Dim resultado As String
    resultado = "Hola" & " " & "Mundo"
    
    ' Esta prueba pasa silenciosamente - no genera errores
End Sub

Public Function Test_FuncionQueDevuelveValor() As String
    ' Prueba como funcion que siempre pasa
    Dim resultado As String
    resultado = "Prueba exitosa"
    
    Test_FuncionQueDevuelveValor = resultado
End Function

Public Sub Test_PruebaQueFalla()
    ' Prueba que falla - genera un error de division por cero
    On Error Resume Next
    Dim resultado As Double
    resultado = 1 / 0  ' Esto genera error 11: Division por cero
    
    ' Verificar si ocurrio un error
    If Err.Number <> 0 Then
        ' El error fue capturado correctamente - la prueba falla como se esperaba
        Err.Clear
    End If
End Sub

Public Sub Test_ValidacionNumero()
    ' Prueba que valida un numero
    Dim numero As Double
    numero = 3.14159
    
    ' Esta prueba pasa silenciosamente - el numero esta en rango
End Sub