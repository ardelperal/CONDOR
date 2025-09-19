' ===== UI-AS-CODE CORE HELPERS =====

Sub UiExportForm(app, dbPath, formName, outputPath, pretty)
    Dim data, source, payload, moduleInfo
    Set data = UiCreateDictionary()
    data("schemaVersion") = 1
    data("name") = formName
    data("objectType") = "acForm"
    data("generatedAtUTC") = UiTimestampUTC()

    Set source = UiCreateDictionary()
    source("database") = dbPath
    source("tool") = "condor_cli.vbs"
    data("source") = source

    Set payload = UiCreateDictionary()
    payload("format") = "AccessText"
    payload("encoding") = "utf-16le"
    payload("text") = UiSerializeForm(app, formName)
    data("payload") = payload

    Set moduleInfo = UiCreateDictionary()
    moduleInfo("name") = formName
    moduleInfo("content") = UiExtractModuleCode(app, formName)
    data("module") = moduleInfo

    Call UiEnsureParentFolder(outputPath)
    Call UiWriteUtf8(outputPath, JsonStringify(data, pretty))
End Sub

Function UiSerializeForm(app, formName)
    Dim tempPath
    tempPath = UiTempPath("condor_export_", ".txt")
    On Error Resume Next
    If objFSO.FileExists(tempPath) Then objFSO.DeleteFile tempPath, True
    On Error GoTo 0
    app.SaveAsText acForm, formName, tempPath
    UiSerializeForm = UiReadUtf16(tempPath)
    On Error Resume Next
    If objFSO.FileExists(tempPath) Then objFSO.DeleteFile tempPath, True
    On Error GoTo 0
End Function

Function UiExtractModuleCode(app, formName)
    On Error Resume Next
    Dim code, formObj, moduleObj
    code = ""
    app.DoCmd.OpenForm formName, acDesign
    If Err.Number = 0 Then
        Set formObj = app.Forms(formName)
        If formObj.HasModule Then
            Set moduleObj = formObj.Module
            If moduleObj.CountOfLines > 0 Then
                code = moduleObj.Lines(1, moduleObj.CountOfLines)
            End If
        End If
        app.DoCmd.Close acForm, formName, acSaveNo
    Else
        Err.Clear
    End If
    On Error GoTo 0
    UiExtractModuleCode = code
End Function

Function UiLoadFormJson(jsonPath)
    UiLoadFormJson = JsonParse(UiReadUtf8(jsonPath))
End Function

Function UiValidateFormJson(data, strict, message)
    message = ""
    If TypeName(data) <> "Dictionary" Then
        message = "La raíz debe ser un objeto JSON"
        UiValidateFormJson = False
        Exit Function
    End If

    If Not data.Exists("schemaVersion") Then
        message = "Falta 'schemaVersion'"
        UiValidateFormJson = False
        Exit Function
    End If
    If Not data.Exists("payload") Then
        message = "Falta bloque 'payload'"
        UiValidateFormJson = False
        Exit Function
    End If

    Dim payload
    Set payload = data("payload")
    If TypeName(payload) <> "Dictionary" Then
        message = "'payload' debe ser un objeto"
        UiValidateFormJson = False
        Exit Function
    End If
    If Not payload.Exists("text") Then
        message = "Falta 'payload.text'"
        UiValidateFormJson = False
        Exit Function
    End If

    If strict Then
        If Not data.Exists("name") Or Len(Trim(CStr(data("name")))) = 0 Then
            message = "Falta 'name' en modo estricto"
            UiValidateFormJson = False
            Exit Function
        End If
    End If

    UiValidateFormJson = True

Function UiGetJsonString(data, key)
    If TypeName(data) = "Dictionary" And data.Exists(key) Then
        UiGetJsonString = CStr(data(key))
    Else
        UiGetJsonString = ""
    End If
End Function

Sub UiImportForm(app, data, targetName, replaceExisting)
    Dim payload
    Set payload = data("payload")

    If replaceExisting Then
        On Error Resume Next
        app.DoCmd.Close acForm, targetName, acSaveNo
        Err.Clear
        app.DoCmd.DeleteObject acForm, targetName
        On Error GoTo 0
    End If

    Dim tempPath
    tempPath = UiTempPath("condor_import_", ".txt")
    Call UiWriteUtf16(tempPath, payload("text"))
    app.LoadFromText acForm, targetName, tempPath
    On Error Resume Next
    If objFSO.FileExists(tempPath) Then objFSO.DeleteFile tempPath, True
    On Error GoTo 0

    If data.Exists("module") Then
        Dim moduleInfo
        Set moduleInfo = data("module")
        If TypeName(moduleInfo) = "Dictionary" And moduleInfo.Exists("content") Then
            Call UiApplyModuleCode(app, targetName, moduleInfo("content"))
        End If
    End If
End Sub

Sub UiApplyModuleCode(app, formName, moduleCode)
    On Error Resume Next
    app.DoCmd.OpenForm formName, acDesign
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    Dim formObj
    Set formObj = app.Forms(formName)
    If Not formObj.HasModule Then formObj.HasModule = True
    Dim moduleObj
    Set moduleObj = formObj.Module
    If moduleObj.CountOfLines > 0 Then moduleObj.DeleteLines 1, moduleObj.CountOfLines
    If Len(moduleCode) > 0 Then moduleObj.InsertLines 1, moduleCode
    app.DoCmd.Close acForm, formName, acSaveYes
    On Error GoTo 0
End Sub

Function UiPerformRoundtrip(dbPath, formName, tempDir, password, pretty)
    Dim result
    Set result = UiCreateDictionary()
    result("success") = False

    Dim prePath, postPath
    prePath = objFSO.BuildPath(tempDir, formName & ".json")
    postPath = objFSO.BuildPath(tempDir, formName & ".post.json")

    Dim app
    Set app = OpenAccessQuiet(dbPath, password)
    If app Is Nothing Then
        Err.Raise vbObjectError + 3501, "UiPerformRoundtrip", "No se pudo abrir Access"
    End If

    Call UiExportForm(app, dbPath, formName, prePath, pretty)
    Call UiImportForm(app, UiLoadFormJson(prePath), formName, True)
    Call UiExportForm(app, dbPath, formName, postPath, pretty)

    Call CloseAccessQuiet(app)

    Dim preData, postData
    Set preData = UiLoadFormJson(prePath)
    Set postData = UiLoadFormJson(postPath)
    Call UiNormalizeFormData(preData)
    Call UiNormalizeFormData(postData)

    If JsonStringify(preData, False) = JsonStringify(postData, False) Then
        result("success") = True
    End If
    result("prePath") = prePath
    result("postPath") = postPath
    Set UiPerformRoundtrip = result
End Function

Sub UiNormalizeFormData(data)
    On Error Resume Next
    If TypeName(data) <> "Dictionary" Then Exit Sub
    If data.Exists("generatedAtUTC") Then data.Remove "generatedAtUTC"
    If data.Exists("source") Then
        Dim src
        Set src = data("source")
        If TypeName(src) = "Dictionary" Then
            If src.Exists("timestamp") Then src.Remove "timestamp"
        End If
    End If
    On Error GoTo 0
End Sub

Sub UiPrintFormJsonSchema(version)
    WScript.Echo "Esquema JSON v" & version & " (UI as Code):"
    WScript.Echo "{"
    WScript.Echo "  ""schemaVersion"": <number>,"
    WScript.Echo "  ""name"": <string>,"
    WScript.Echo "  ""objectType"": \"acForm\","
    WScript.Echo "  ""generatedAtUTC"": <string>,"
    WScript.Echo "  ""source"": { ... },"
    WScript.Echo "  ""payload"": { "format": \"AccessText\", "text": <string> },"
    WScript.Echo "  ""module"": { "name": <string>, "content": <string> }"
    WScript.Echo "}"
End Sub

Function UiCreateDictionary()
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = TextCompare
    Set UiCreateDictionary = d
End Function

Function UiCreateArrayList()
    Set UiCreateArrayList = CreateObject("System.Collections.ArrayList")
End Function

Sub UiEnsureParentFolder(path)
    Dim parent
    parent = objFSO.GetParentFolderName(path)
    If parent <> "" And Not objFSO.FolderExists(parent) Then
        objFSO.CreateFolder parent
    End If
End Sub

Function UiTimestampUTC()
    Dim dt
    dt = Now
    UiTimestampUTC = Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2) & _
        "T" & Right("0" & Hour(dt), 2) & ":" & Right("0" & Minute(dt), 2) & ":" & Right("0" & Second(dt), 2) & "Z"
End Function

Function UiTempPath(prefix, suffix)
    Dim tempFolder
    tempFolder = objFSO.GetSpecialFolder(2)
    UiTempPath = objFSO.BuildPath(tempFolder, prefix & Hex(Timer * 1000) & suffix)
End Function
