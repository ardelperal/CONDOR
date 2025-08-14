' Script para importar el módulo de prueba a Access
Dim accessApp
Set accessApp = CreateObject("Access.Application")

On Error Resume Next

' Abrir la base de datos
accessApp.OpenCurrentDatabase "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"

If Err.Number <> 0 Then
    WScript.Echo "Error al abrir la base de datos: " & Err.Description
    WScript.Quit 1
End If

Err.Clear

' Importar el módulo de prueba
accessApp.DoCmd.TransferText acImportDelim, , "TestCompilation", "C:\Proyectos\CONDOR\test_compilation.bas", False

If Err.Number <> 0 Then
    ' Intentar con LoadFromText
    Err.Clear
    accessApp.LoadFromText acModule, "TestCompilation", "C:\Proyectos\CONDOR\test_compilation.bas"
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al importar el módulo: " & Err.Description
        WScript.Echo "Número de error: " & Err.Number
    Else
        WScript.Echo "Módulo TestCompilation importado exitosamente con LoadFromText"
    End If
Else
    WScript.Echo "Módulo TestCompilation importado exitosamente con TransferText"
End If

' Cerrar Access
accessApp.Quit
Set accessApp = Nothing

WScript.Echo "Proceso completado. Ahora puedes ejecutar TestCompilation() desde Access."