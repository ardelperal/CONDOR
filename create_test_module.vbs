' Script para crear módulo de prueba directamente en Access
Dim accessApp
Set accessApp = CreateObject("Access.Application")
accessApp.Visible = True

On Error Resume Next

' Abrir la base de datos
accessApp.OpenCurrentDatabase "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"

If Err.Number <> 0 Then
    WScript.Echo "Error al abrir la base de datos: " & Err.Description
    WScript.Quit 1
End If

Err.Clear

' Crear el módulo directamente
Dim vbProj, newModule
Set vbProj = accessApp.VBE.ActiveVBProject
Set newModule = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
newModule.Name = "TestCompilation"

' Agregar el código al módulo
Dim moduleCode
moduleCode = "Option Compare Database" & vbCrLf & _
             "Option Explicit" & vbCrLf & vbCrLf & _
             "Public Function TestCompilation() As Boolean" & vbCrLf & _
             "    On Error GoTo ErrorHandler" & vbCrLf & vbCrLf & _
             "    ' Intentar crear una instancia de CSolicitudPC" & vbCrLf & _
             "    Dim solicitud As ISolicitud" & vbCrLf & _
             "    Set solicitud = New CSolicitudPC" & vbCrLf & vbCrLf & _
             "    ' Si llegamos aquí, la compilación es exitosa" & vbCrLf & _
             "    TestCompilation = True" & vbCrLf & _
             "    Debug.Print ""Compilación exitosa - CSolicitudPC creado correctamente""" & vbCrLf & vbCrLf & _
             "    Exit Function" & vbCrLf & vbCrLf & _
             "ErrorHandler:" & vbCrLf & _
             "    TestCompilation = False" & vbCrLf & _
             "    Debug.Print ""Error de compilación: "" & Err.Description" & vbCrLf & _
             "    Debug.Print ""Número de error: "" & Err.Number" & vbCrLf & _
             "End Function"

newModule.CodeModule.AddFromString moduleCode

If Err.Number <> 0 Then
    WScript.Echo "Error al crear el módulo: " & Err.Description
    WScript.Echo "Número de error: " & Err.Number
Else
    WScript.Echo "Módulo TestCompilation creado exitosamente"
    WScript.Echo "Ahora puedes ejecutar TestCompilation() desde la ventana Inmediato (Ctrl+G)"
End If

WScript.Echo "Access permanece abierto para que puedas probar la función"