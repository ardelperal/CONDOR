' Test script para verificar la implementación de CONDOR
' Este script prueba todas las funciones implementadas

Option Explicit

' Incluir el archivo principal
Dim fso, scriptPath, mainScript, fileContent
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
mainScript = fso.BuildPath(scriptPath, "condor_cli.vbs")

' Ejecutar el archivo principal para cargar las funciones
fileContent = fso.OpenTextFile(mainScript).ReadAll()
ExecuteGlobal fileContent

' Variables de prueba
Dim testResults
testResults = ""

Sub AddTestResult(testName, result, expected, actual)
    testResults = testResults & testName & ": " & result & vbCrLf
    If result = "FAIL" Then
        testResults = testResults & "  Esperado: " & expected & vbCrLf
        testResults = testResults & "  Actual: " & actual & vbCrLf
    End If
    testResults = testResults & vbCrLf
End Sub

' Test 1: MapPropKey
Sub TestMapPropKey()
    Dim result1, result2, result3
    result1 = MapPropKey("ancho")
    result2 = MapPropKey("width")
    result3 = MapPropKey("unknownprop")
    
    If result1 = "Width" Then
        AddTestResult "MapPropKey(ancho)", "PASS", "Width", result1
    Else
        AddTestResult "MapPropKey(ancho)", "FAIL", "Width", result1
    End If
    
    If result2 = "Width" Then
        AddTestResult "MapPropKey(width)", "PASS", "Width", result2
    Else
        AddTestResult "MapPropKey(width)", "FAIL", "Width", result2
    End If
    
    If result3 = "unknownprop" Then
        AddTestResult "MapPropKey(unknownprop)", "PASS", "unknownprop", result3
    Else
        AddTestResult "MapPropKey(unknownprop)", "FAIL", "unknownprop", result3
    End If
End Sub

' Test 2: NormalizeEnumToken
Sub TestNormalizeEnumToken()
    Dim result1, result2, result3
    result1 = NormalizeEnumToken("verdadero", "Boolean")
    result2 = NormalizeEnumToken("true", "Boolean")
    result3 = NormalizeEnumToken("unknown", "Boolean")
    
    If result1 = "True" Then
        AddTestResult "NormalizeEnumToken(verdadero)", "PASS", "True", result1
    Else
        AddTestResult "NormalizeEnumToken(verdadero)", "FAIL", "True", result1
    End If
    
    If result2 = "True" Then
        AddTestResult "NormalizeEnumToken(true)", "PASS", "True", result2
    Else
        AddTestResult "NormalizeEnumToken(true)", "FAIL", "True", result2
    End If
    
    If result3 = "unknown" Then
        AddTestResult "NormalizeEnumToken(unknown)", "PASS", "unknown", result3
    Else
        AddTestResult "NormalizeEnumToken(unknown)", "FAIL", "unknown", result3
    End If
End Sub

' Test 3: ConvertColorToLong
Sub TestConvertColorToLong()
    Dim result1, result2
    result1 = ConvertColorToLong("#FF0000")  ' Rojo
    result2 = ConvertColorToLong("#00FF00")  ' Verde
    
    If result1 = 255 Then  ' BGR: 0x0000FF
        AddTestResult "ConvertColorToLong(#FF0000)", "PASS", "255", result1
    Else
        AddTestResult "ConvertColorToLong(#FF0000)", "FAIL", "255", result1
    End If
    
    If result2 = 65280 Then  ' BGR: 0x00FF00
        AddTestResult "ConvertColorToLong(#00FF00)", "PASS", "65280", result2
    Else
        AddTestResult "ConvertColorToLong(#00FF00)", "FAIL", "65280", result2
    End If
End Sub

' Test 4: MapEventToProperty
Sub TestMapEventToProperty()
    Dim result1, result2, result3
    result1 = MapEventToProperty("clic")
    result2 = MapEventToProperty("click")
    result3 = MapEventToProperty("unknown")
    
    If result1 = "OnClick" Then
        AddTestResult "MapEventToProperty(clic)", "PASS", "OnClick", result1
    Else
        AddTestResult "MapEventToProperty(clic)", "FAIL", "OnClick", result1
    End If
    
    If result2 = "OnClick" Then
        AddTestResult "MapEventToProperty(click)", "PASS", "OnClick", result2
    Else
        AddTestResult "MapEventToProperty(click)", "FAIL", "OnClick", result2
    End If
    
    If result3 = "" Then
        AddTestResult "MapEventToProperty(unknown)", "PASS", """""", result3
    Else
        AddTestResult "MapEventToProperty(unknown)", "FAIL", """""", result3
    End If
End Sub

' Test 5: ParseJsonObject
Sub TestParseJsonObject()
    Dim jsonText, result
    jsonText = "{""name"":""TestForm"",""width"":800,""height"":600}"
    
    On Error Resume Next
    Set result = ParseJsonObject(jsonText)
    
    If Err.Number = 0 And Not result Is Nothing Then
        If result.Exists("name") And result("name") = "TestForm" Then
            AddTestResult "ParseJsonObject(basic)", "PASS", "Object with name=TestForm", "Object created successfully"
        Else
            AddTestResult "ParseJsonObject(basic)", "FAIL", "Object with name=TestForm", "Object missing name property"
        End If
    Else
        AddTestResult "ParseJsonObject(basic)", "FAIL", "Valid object", "Error: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' Ejecutar todas las pruebas
Sub RunAllTests()
    testResults = "=== RESULTADOS DE PRUEBAS CONDOR ===" & vbCrLf & vbCrLf
    
    TestMapPropKey
    TestNormalizeEnumToken
    TestConvertColorToLong
    TestMapEventToProperty
    TestParseJsonObject
    
    ' Mostrar resultados
    WScript.Echo testResults
    
    ' Guardar resultados en archivo
    Dim outputFile
    outputFile = fso.BuildPath(scriptPath, "test_results.txt")
    fso.OpenTextFile(outputFile, 2, True).Write testResults
    WScript.Echo "Resultados guardados en: " & outputFile
End Sub

' Ejecutar las pruebas
RunAllTests