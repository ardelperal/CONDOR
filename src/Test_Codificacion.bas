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

