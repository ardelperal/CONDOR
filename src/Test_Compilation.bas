Attribute VB_Name = "Test_Compilation"
Option Compare Database
Option Explicit

' M?dulo de prueba para verificar compilaci?n

Public Function TestCompilation() As Boolean
    On Error GoTo ErrorHandler
    
    ' Intentar crear una instancia de CSolicitudPC
    Dim solicitud As ISolicitud
    Set solicitud = New CSolicitudPC
    
    ' Si llegamos aqu?, la compilaci?n es exitosa
    TestCompilation = True
    Debug.Print "Compilaci?n exitosa - CSolicitudPC creado correctamente"
    
    Exit Function
    
ErrorHandler:
    TestCompilation = False
    Debug.Print "Error de compilaci?n: " & Err.Description
    Debug.Print "N?mero de error: " & Err.Number
End Function



