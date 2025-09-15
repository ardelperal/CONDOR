Option Explicit

Dim objAccess, objDB, objForm, objFSO, objFile
Dim strDbPath, strPassword, strFormName, strOutputPath

' Configuración
strDbPath = "C:\Proyectos\CONDOR\ui\sources\Expedientes.accdb"
strPassword = "dpddpd"
strFormName = "Formulario1"
strOutputPath = "C:\Proyectos\CONDOR\.tmp\Formulario1_info.json"

Set objFSO = CreateObject("Scripting.FileSystemObject")

On Error Resume Next

' Crear instancia de Access
Set objAccess = CreateObject("Access.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error creando Access: " & Err.Description
    WScript.Quit 1
End If

' Abrir base de datos
objAccess.OpenCurrentDatabase strDbPath, False, strPassword
If Err.Number <> 0 Then
    WScript.Echo "Error abriendo base de datos: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If

' Abrir formulario en modo diseño
objAccess.DoCmd.OpenForm strFormName, 0 ' acDesign = 0
If Err.Number <> 0 Then
    WScript.Echo "Error abriendo formulario: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If

' Obtener referencia al formulario
Set objForm = objAccess.Forms(strFormName)
If Err.Number <> 0 Then
    WScript.Echo "Error obteniendo referencia al formulario: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If

' Crear archivo de salida
Set objFile = objFSO.CreateTextFile(strOutputPath, True)

' Escribir información básica del formulario
objFile.WriteLine "{"
objFile.WriteLine "  ""schemaVersion"": ""1.0.0"","
objFile.WriteLine "  ""units"": ""twips"","
objFile.WriteLine "  ""formName"": """ & strFormName & ""","
objFile.WriteLine "  ""properties"": {"
objFile.WriteLine "    ""caption"": """ & Replace(objForm.Caption, """", """""") & ""","
objFile.WriteLine "    ""width"": " & objForm.Width & ","
objFile.WriteLine "    ""height"": " & objForm.WindowHeight
objFile.WriteLine "  },"
objFile.WriteLine "  ""sections"": {"

' Información de la sección Detail
objFile.WriteLine "    ""Detail"": {"
objFile.WriteLine "      ""properties"": {"
objFile.WriteLine "        ""height"": " & objForm.Section(0).Height
objFile.WriteLine "      },"
objFile.WriteLine "      ""controls"": ["

' Listar controles básicos
Dim i, ctrl, bFirst
bFirst = True
For i = 0 To objForm.Controls.Count - 1
    Set ctrl = objForm.Controls(i)
    If Not bFirst Then objFile.WriteLine ","
    objFile.WriteLine "        {"
    objFile.WriteLine "          ""name"": """ & ctrl.Name & ""","
    objFile.WriteLine "          ""type"": """ & TypeName(ctrl) & ""","
    objFile.WriteLine "          ""properties"": {"
    objFile.WriteLine "            ""top"": " & ctrl.Top & ","
    objFile.WriteLine "            ""left"": " & ctrl.Left & ","
    objFile.WriteLine "            ""width"": " & ctrl.Width & ","
    objFile.WriteLine "            ""height"": " & ctrl.Height
    If ctrl.Caption <> "" Then
        objFile.WriteLine "            ,""caption"": """ & Replace(ctrl.Caption, """", """""") & """"
    End If
    objFile.WriteLine "          }"
    objFile.WriteLine "        }"
    bFirst = False
Next

objFile.WriteLine "      ]"
objFile.WriteLine "    }"
objFile.WriteLine "  }"
objFile.WriteLine "}"

objFile.Close

' Cerrar Access
objAccess.Quit

WScript.Echo "Información del formulario extraída exitosamente en: " & strOutputPath

On Error GoTo 0