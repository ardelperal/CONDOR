' Script para ejecutar todas las pruebas de CONDOR desde línea de comandos
' Uso: cscript run_tests.vbs

Option Explicit

Dim objAccess
Dim strAccessPath
Dim objFSO

' Configuración
strAccessPath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"

' Verificar que existe la base de datos
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FileExists(strAccessPath) Then
    WScript.Echo "Error: No se encuentra la base de datos en " & strAccessPath
    WScript.Quit 1
End If

' Crear instancia de Access
On Error Resume Next
Set objAccess = CreateObject("Access.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: No se puede crear instancia de Access: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

' Abrir la base de datos
On Error Resume Next
objAccess.OpenCurrentDatabase strAccessPath
If Err.Number <> 0 Then
    WScript.Echo "Error: No se puede abrir la base de datos: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "=== EJECUTANDO PRUEBAS DE CONDOR ==="
WScript.Echo "Base de datos: " & strAccessPath
WScript.Echo "Fecha y hora: " & Now()
WScript.Echo ""

' Ejecutar las pruebas
On Error Resume Next
Dim resultado
resultado = objAccess.Run("EJECUTAR_TODAS_LAS_PRUEBAS")
If Err.Number <> 0 Then
    WScript.Echo "Error ejecutando pruebas: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "=== PRUEBAS COMPLETADAS ==="

' Cerrar Access
objAccess.Quit
Set objAccess = Nothing
Set objFSO = Nothing

WScript.Echo "Nota: Los resultados detallados se muestran en la Ventana Inmediato de Access."
WScript.Echo "Para ver los resultados, abra Access y presione Ctrl+G."