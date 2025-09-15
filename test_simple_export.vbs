Option Explicit

' Script de prueba simple para ExportFormInternal
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Incluir las constantes necesarias
Const acViewDesign = 1
Const acDetail = 0
Const acForm = 2
Const acSaveNo = 2

' Función helper para convertir OLE a Hex
Function OleToHex(ole)
    Dim r, g, b
    r = (ole And &HFF)
    g = (ole \ &H100) And &HFF
    b = (ole \ &H10000) And &HFF
    OleToHex = "#" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
End Function

' Función OpenAccessApp simplificada
Function OpenAccessApp(dbPath, password, bypassStartup)
    On Error Resume Next
    Dim app
    Set app = CreateObject("Access.Application")
    If Err.Number <> 0 Then
        WScript.Echo "Error creando Access.Application: " & Err.Description
        Set OpenAccessApp = Nothing
        Exit Function
    End If
    
    app.Visible = False
    app.UserControl = False
    app.AutomationSecurity = 3
    
    If password <> "" Then
        app.OpenCurrentDatabase dbPath, False, password
    Else
        app.OpenCurrentDatabase dbPath, False
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "Error abriendo base de datos: " & Err.Description
        app.Quit
        Set app = Nothing
        Set OpenAccessApp = Nothing
        Exit Function
    End If
    
    app.Echo False
    Set OpenAccessApp = app
    On Error GoTo 0
End Function

' Función CloseAccessApp simplificada
Sub CloseAccessApp(app)
    On Error Resume Next
    If Not app Is Nothing Then
        app.Echo True
        app.CloseCurrentDatabase
        app.Quit
        Set app = Nothing
    End If
End Sub

' Función ExportFormInternal simplificada para prueba
Sub ExportFormInternal(dbPath, formName, outputPath, password)
    WScript.Echo "=== PRUEBA EXPORTFORMINTERNAL ==="
    WScript.Echo "DB: " & dbPath
    WScript.Echo "Form: " & formName
    WScript.Echo "Output: " & outputPath
    
    ' Crear instancia de Access
    Dim objAccessLocal
    Set objAccessLocal = OpenAccessApp(dbPath, password, True)
    
    If objAccessLocal Is Nothing Then
        WScript.Echo "ERROR: No se pudo abrir la base de datos"
        Exit Sub
    End If
    
    On Error Resume Next
    
    ' Intentar abrir formulario en vista Diseño
    WScript.Echo "Intentando abrir formulario en vista Diseño..."
    objAccessLocal.DoCmd.OpenForm formName, acViewDesign
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo abrir formulario: " & Err.Description
        CloseAccessApp objAccessLocal
        Exit Sub
    End If
    
    WScript.Echo "Formulario abierto exitosamente"
    
    ' Obtener referencia al formulario
    Dim frm
    Set frm = objAccessLocal.Forms(formName)
    
    ' Leer propiedades básicas
    Dim caption, width, height, backColor, recordSource, recordsetType
    
    caption = frm.Caption
    width = frm.Width
    height = frm.Section(acDetail).Height
    backColor = OleToHex(frm.Section(acDetail).BackColor)
    recordSource = frm.RecordSource
    recordsetType = frm.RecordsetType
    
    WScript.Echo "Propiedades leídas:"
    WScript.Echo "  Caption: " & caption
    WScript.Echo "  Width: " & width
    WScript.Echo "  Height: " & height
    WScript.Echo "  BackColor: " & backColor
    WScript.Echo "  RecordSource: " & recordSource
    WScript.Echo "  RecordsetType: " & recordsetType
    
    ' Generar JSON simple
    Dim jsonContent
    jsonContent = "{" & vbCrLf
    jsonContent = jsonContent & "  ""formName"": """ & frm.Name & """," & vbCrLf
    jsonContent = jsonContent & "  ""caption"": """ & caption & """," & vbCrLf
    jsonContent = jsonContent & "  ""width"": " & width & "," & vbCrLf
    jsonContent = jsonContent & "  ""height"": " & height & vbCrLf
    jsonContent = jsonContent & "}" & vbCrLf
    
    ' Cerrar formulario
    objAccessLocal.DoCmd.Close acForm, formName, acSaveNo
    
    ' Guardar archivo
    Dim objFile
    Set objFile = objFSO.CreateTextFile(outputPath, True, True)
    objFile.Write jsonContent
    objFile.Close
    
    WScript.Echo "JSON guardado en: " & outputPath
    
    CloseAccessApp objAccessLocal
    WScript.Echo "=== PRUEBA COMPLETADA ==="
    On Error GoTo 0
End Sub

' Ejecutar prueba con ruta absoluta
Dim currentDir
currentDir = objFSO.GetAbsolutePathName(".")
Dim dbFullPath
dbFullPath = objFSO.BuildPath(currentDir, "ui\sources\Expedientes.accdb")

WScript.Echo "Ruta completa de BD: " & dbFullPath
WScript.Echo "Archivo existe: " & objFSO.FileExists(dbFullPath)

Call ExportFormInternal(dbFullPath, "Formulario1", ".\test_export.json", "dpddpd")