
' ===== FORM COMMANDS (UI-AS-CODE) =====

Sub ExportFormCommand()
    On Error GoTo ExportErr

    Dim dbPath, formName, outputPath, password, pretty, expand
    Dim i, arg

    dbPath = ""
    formName = ""
    outputPath = ""
    password = gPassword
    pretty = False
    expand = ""

    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        If Left(arg, 2) = "--" Then
            Select Case LCase(arg)
                Case "--output"
                    If i < objArgs.Count - 1 Then
                        outputPath = TrimQuotes(objArgs(i + 1))
                        i = i + 1
                    End If
                Case "--password"
                    If i < objArgs.Count - 1 Then
                        password = objArgs(i + 1)
                        i = i + 1
                    End If
                Case "--pretty"
                    pretty = True
                Case "--expand"
                    If i < objArgs.Count - 1 Then
                        expand = LCase(objArgs(i + 1))
                        i = i + 1
                    End If
                Case "--help"
                    Call ShowExportFormHelp()
                    WScript.Quit 0
            End Select
        ElseIf Left(arg, 1) <> "-" Then
            If dbPath = "" Then
                dbPath = arg
            ElseIf formName = "" Then
                formName = arg
            End If
        End If
    Next

    If dbPath = "" Then
        dbPath = strAccessPath
    End If

    If formName = "" Then
        WScript.Echo "[ERROR] Falta el nombre del formulario."
        Call ShowExportFormHelp()
        WScript.Quit 1
    End If

    If outputPath = "" Then
        WScript.Echo "[ERROR] Debe indicar --output <ruta.json>."
        Call ShowExportFormHelp()
        WScript.Quit 1
    End If

    dbPath = ToAbsolute(dbPath)
    outputPath = ToAbsolute(outputPath)

    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "[ERROR] Base de datos no encontrada: " & dbPath
        WScript.Quit 1
    End If

    Dim app
    Set app = OpenAccessQuiet(dbPath, password)
    If app Is Nothing Then
        WScript.Echo "[ERROR] No se pudo abrir Access en " & dbPath
        WScript.Quit 1
    End If

    Call ExportFormToJsonFile(app, dbPath, formName, outputPath, pretty)

    Call CloseAccessQuiet(app)

    WScript.Echo "Formulario '" & formName & "' exportado a " & outputPath
    WScript.Quit 0

ExportErr:
    Dim errMsg
    errMsg = "[ERROR] " & Err.Description
    On Error Resume Next
    If Not app Is Nothing Then
        Call CloseAccessQuiet(app)
    End If
    WScript.Echo errMsg
    WScript.Quit 1
End Sub

Sub ImportFormCommand()
    On Error GoTo ImportErr

    Dim dbPath, jsonPath, targetName, password, replaceExisting, strict
    Dim i, arg

    dbPath = ""
    jsonPath = ""
    targetName = ""
    password = gPassword
    replaceExisting = False
    strict = False

    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        If Left(arg, 2) = "--" Then
            Select Case LCase(arg)
                Case "--target"
                    If i < objArgs.Count - 1 Then
                        targetName = objArgs(i + 1)
                        i = i + 1
                    End If
                Case "--password"
                    If i < objArgs.Count - 1 Then
                        password = objArgs(i + 1)
                        i = i + 1
                    End If
                Case "--replace"
                    replaceExisting = True
                Case "--strict"
                    strict = True
                Case "--help"
                    Call ShowImportFormHelp()
                    WScript.Quit 0
            End Select
        ElseIf Left(arg, 1) <> "-" Then
            If dbPath = "" Then
                dbPath = arg
            ElseIf jsonPath = "" Then
                jsonPath = arg
            End If
        End If
    Next

    If jsonPath = "" Then
        WScript.Echo "[ERROR] Debe indicar el archivo JSON de entrada."
        Call ShowImportFormHelp()
        WScript.Quit 1
    End If

    If dbPath = "" Then
        dbPath = strAccessPath
    End If

    dbPath = ToAbsolute(dbPath)
    jsonPath = ToAbsolute(jsonPath)

    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "[ERROR] Base de datos no encontrada: " & dbPath
        WScript.Quit 1
    End If

    If Not objFSO.FileExists(jsonPath) Then
        WScript.Echo "[ERROR] Archivo JSON no encontrado: " & jsonPath
        WScript.Quit 1
    End If

    Dim data
    Set data = JsonParseFile(jsonPath)

    Dim validationMessage
    If Not ValidateFormJsonStructure(data, strict, validationMessage) Then
        WScript.Echo "[ERROR] " & validationMessage
        WScript.Quit 1
    End If

    If targetName = "" Then
        targetName = GetJsonString(data, "name")
    End If

    Dim app
    Set app = OpenAccessQuiet(dbPath, password)
    If app Is Nothing Then
        WScript.Echo "[ERROR] No se pudo abrir Access en " & dbPath
        WScript.Quit 1
    End If

    Call ImportFormFromJsonData(app, data, targetName, replaceExisting)

    Call CloseAccessQuiet(app)

    WScript.Echo "Formulario importado correctamente: " & targetName
    WScript.Quit 0

ImportErr:
    Dim errMsg
    errMsg = "[ERROR] " & Err.Description
    On Error Resume Next
    If Not app Is Nothing Then
        Call CloseAccessQuiet(app)
    End If
    WScript.Echo errMsg
    WScript.Quit 1
End Sub

