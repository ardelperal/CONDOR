' Script VBS para probar compilación de Access
Dim accessApp
Set accessApp = CreateObject("Access.Application")

On Error Resume Next

' Abrir la base de datos
accessApp.OpenCurrentDatabase "C:\Proyectos\CONDOR\CONDOR.accdb"

If Err.Number <> 0 Then
    WScript.Echo "Error al abrir la base de datos: " & Err.Description
    WScript.Quit 1
End If

' Intentar compilar
accessApp.DoCmd.RunCommand 10

If Err.Number <> 0 Then
    WScript.Echo "Error de compilación: " & Err.Description
    WScript.Echo "Número de error: " & Err.Number
Else
    WScript.Echo "Compilación exitosa"
End If

' Cerrar Access
accessApp.Quit
Set accessApp = Nothing