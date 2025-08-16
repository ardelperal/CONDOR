Option Explicit

' ======================
' CLI Access Form IO Tool
' ======================
'
' Uso:
'   cscript //nologo cli.vbs exportar "C:\MiBase.accdb" "C:\Salida.json" [NombreFormulario]
'   cscript //nologo cli.vbs importar "C:\MiBase.accdb" "C:\Entrada.json"
'
' ======================

Dim args, modo, rutaBase, rutaJSON, nombreFormulario
Set args = WScript.Arguments

If args.Count < 3 Then
    WScript.Echo "Uso:"
    WScript.Echo "  Exportar: cscript //nologo cli.vbs exportar <rutaBase> <rutaJSON> [nombreFormulario]"
    WScript.Echo "  Importar: cscript //nologo cli.vbs importar <rutaBase> <rutaJSON>"
    WScript.Quit 1
End If

modo = LCase(args(0))
rutaBase = args(1)
rutaJSON = args(2)
If args.Count >= 4 Then nombreFormulario = args(3)

Dim accessApp
Set accessApp = CreateObject("Access.Application")
accessApp.OpenCurrentDatabase rutaBase

Select Case modo
    Case "exportar"
        ExportarFormularios accessApp, rutaJSON, nombreFormulario
    Case "importar"
        ImportarFormularios accessApp, rutaJSON
    Case Else
        WScript.Echo "Modo no reconocido: " & modo
End Select

accessApp.Quit
Set accessApp = Nothing
WScript.Echo "Proceso completado."

' ======================
' Función: ExportarFormularios
' ======================
Sub ExportarFormularios(acc, salidaJSON, optNombre)
    Dim fName, formObj, ctl, json, fso, ts, propValue

    json = "{""formularios"":["

    For Each fName In acc.CurrentProject.AllForms
        If optNombre <> "" And fName.Name <> optNombre Then
            ' Si se pasó un nombre y este no coincide, saltar
            Continue For
        End If

        acc.DoCmd.OpenForm fName.Name, 1 ' Vista diseño
        Set formObj = acc.Forms(fName.Name)

        json = json & "{""nombre"":""" & EscapeJSON(fName.Name) & """,""controles"":["

        For Each ctl In formObj.Controls
            json = json & "{"
            json = json & """nombre"":""" & EscapeJSON(ctl.Name) & """"
            json = json & ",""tipo"":" & ctl.ControlType
            json = json & ",""izquierda"":" & ctl.Left
            json = json & ",""arriba"":" & ctl.Top
            json = json & ",""ancho"":" & ctl.Width
            json = json & ",""alto"":" & ctl.Height

            propValue = GetProp(ctl, "Caption")
            If Not IsNull(propValue) Then json = json & ",""caption"":""" & EscapeJSON(propValue) & """"

            propValue = GetProp(ctl, "ControlSource")
            If Not IsNull(propValue) Then json = json & ",""controlSource"":""" & EscapeJSON(propValue) & """"

            propValue = GetProp(ctl, "RowSource")
            If Not IsNull(propValue) Then json = json & ",""rowSource"":""" & EscapeJSON(propValue) & """"

            json = json & "},"
        Next

        If Right(json,1) = "," Then json = Left(json, Len(json)-1)
        json = json & "]},"

        acc.DoCmd.Close 2, fName.Name
    Next

    If Right(json,1) = "," Then json = Left(json, Len(json)-1)
    json = json & "]}"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(salidaJSON, True, False)
    ts.Write json
    ts.Close

    WScript.Echo "Exportación a JSON completada: " & salidaJSON
End Sub

' ======================
' Función: ImportarFormularios
' ======================
Sub ImportarFormularios(acc, entradaJSON)
    Dim fso, ts, jsonText, jsonObj, formObj, ctlData, newCtl, frmName
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(entradaJSON, 1, False)
    jsonText = ts.ReadAll
    ts.Close

    ' Convertir JSON a objeto usando ScriptControl (VBScript no tiene JSON nativo)
    Dim sc
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    Set jsonObj = sc.Eval("(" & jsonText & ")")

    Dim i, j
    For i = 0 To jsonObj.formularios.length - 1
        frmName = jsonObj.formularios(i).nombre
        acc.CreateForm
        acc.DoCmd.Save acForm, frmName
        acc.DoCmd.Close acForm, frmName
        acc.DoCmd.OpenForm frmName, 1 ' Vista diseño
        Set formObj = acc.Forms(frmName)

        For j = 0 To jsonObj.formularios(i).controles.length - 1
            ctlData = jsonObj.formularios(i).controles(j)
            Set newCtl = acc.CreateControl(frmName, ctlData.tipo, , , , ctlData.izquierda, ctlData.arriba, ctlData.ancho, ctlData.alto)
            newCtl.Name = ctlData.nombre
            If Not IsNullJS(ctlData.caption) Then newCtl.Caption = ctlData.caption
            If Not IsNullJS(ctlData.controlSource) Then newCtl.ControlSource = ctlData.controlSource
            If Not IsNullJS(ctlData.rowSource) Then newCtl.RowSource = ctlData.rowSource
        Next

        acc.DoCmd.Save acForm, frmName
        acc.DoCmd.Close acForm, frmName
    Next

    WScript.Echo "Importación desde JSON completada."
End Sub

' ======================
' Funciones auxiliares
' ======================
Function EscapeJSON(str)
    If IsNull(str) Then
        EscapeJSON = ""
        Exit Function
    End If
    Dim tmp
    tmp = Replace(str, "\", "\\")
    tmp = Replace(tmp, """", "\""")
    tmp = Replace(tmp, vbCrLf, "\n")
    tmp = Replace(tmp, vbCr, "\n")
    tmp = Replace(tmp, vbLf, "\n")
    EscapeJSON = tmp
End Function

Function GetProp(obj, propName)
    On Error Resume Next
    GetProp = Null
    Dim val
    val = CallByName(obj, propName, 2)
    If Err.Number = 0 Then GetProp = val
    Err.Clear
    On Error GoTo 0
End Function

Function IsNullJS(val)
    On Error Resume Next
    If IsObject(val) Then
        IsNullJS = False
    Else
        If IsEmpty(val) Then
            IsNullJS = True
        ElseIf IsNull(val) Then
            IsNullJS = True
        Else
            IsNullJS = False
        End If
    End If
    On Error GoTo 0
End Function
