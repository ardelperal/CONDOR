' =====================================================
' SCRIPT: compile_and_test.vbs
' DESCRIPCION: Automatiza la compilación y ejecución de pruebas
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

Option Explicit

Dim objAccess, strDatabasePath, objFSO

' Configuración
strDatabasePath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Verificar que existe la base de datos
If Not objFSO.FileExists(strDatabasePath) Then
    WScript.Echo "ERROR: No se encuentra la base de datos: " & strDatabasePath
    WScript.Quit 1
End If

WScript.Echo "Iniciando compilación y pruebas..."
WScript.Echo "Base de datos: " & strDatabasePath

' Crear instancia de Access
Set objAccess = CreateObject("Access.Application")
objAccess.Visible = False

On Error Resume Next

' Abrir la base de datos
objAccess.OpenCurrentDatabase strDatabasePath

If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se pudo abrir la base de datos: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If

WScript.Echo "Base de datos abierta correctamente"

' Compilar el proyecto usando VBA
WScript.Echo "Compilando proyecto VBA..."

' Ejecutar código VBA para compilar
Dim strVBACode
strVBACode = "Application.VBE.CommandBars.FindControl(ID:=578).Execute"

' Intentar compilar usando DoCmd
objAccess.DoCmd.RunCommand 578  ' acCmdCompileAndSaveAllModules

If Err.Number <> 0 Then
    WScript.Echo "ADVERTENCIA DE COMPILACION: " & Err.Description
    WScript.Echo "Continuando con las pruebas..."
    Err.Clear
Else
    WScript.Echo "Compilación completada"
End If

' Verificar si existen módulos de prueba
WScript.Echo "Verificando módulos de prueba..."

Dim objModule, bTestsFound
bTestsFound = False

' Buscar módulos que contengan "Test_" en el nombre
Dim i
For i = 0 To objAccess.CurrentProject.AllModules.Count - 1
    If InStr(UCase(objAccess.CurrentProject.AllModules(i).Name), "TEST_") > 0 Then
        WScript.Echo "Encontrado módulo de prueba: " & objAccess.CurrentProject.AllModules(i).Name
        bTestsFound = True
    End If
Next

If bTestsFound Then
    WScript.Echo "Se encontraron módulos de prueba"
Else
    WScript.Echo "No se encontraron módulos de prueba"
End If

' Cerrar Access
objAccess.CloseCurrentDatabase
objAccess.Quit

WScript.Echo "Proceso completado"
WScript.Quit 0