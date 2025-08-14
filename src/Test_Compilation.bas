Attribute VB_Name = "Test_Compilation"
' Módulo de prueba para verificar compilación

Public Function TestCompilation() As Boolean
    On Error GoTo ErrorHandler
    
    ' Intentar crear una instancia de CSolicitudPC
    Dim solicitud As ISolicitud
    Set solicitud = New CSolicitudPC
    
    ' Si llegamos aquí, la compilación es exitosa
    TestCompilation = True
    Debug.Print "Compilación exitosa - CSolicitudPC creado correctamente"
    
    Exit Function
    
ErrorHandler:
    TestCompilation = False
    Debug.Print "Error de compilación: " & Err.Description
    Debug.Print "Número de error: " & Err.Number
End Function
