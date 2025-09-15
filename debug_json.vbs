Option Explicit

' Incluir la clase JsonParser del archivo principal
Dim objFSO, objFile, strContent, objParser, result

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Leer el archivo JSON
Set objFile = objFSO.OpenTextFile("test_clean.json", 1, False, 0)
strContent = objFile.ReadAll
objFile.Close

WScript.Echo "Contenido leído: " & strContent
WScript.Echo "Longitud: " & Len(strContent)
WScript.Echo "Primer carácter ASCII: " & Asc(Left(strContent, 1))

' Probar parsing simple
On Error Resume Next
Set objParser = New JsonParser
Set result = objParser.Parse(strContent)

If Err.Number <> 0 Then
    WScript.Echo "Error en parsing: " & Err.Number & " - " & Err.Description
    WScript.Echo "Source: " & Err.Source
Else
    WScript.Echo "Parsing exitoso"
    WScript.Echo "Tipo de resultado: " & TypeName(result)
    If Not result Is Nothing Then
        WScript.Echo "Resultado no es Nothing"
        If TypeName(result) = "Dictionary" Then
            WScript.Echo "Es un Dictionary con " & result.Count & " elementos"
        End If
    Else
        WScript.Echo "Resultado es Nothing"
    End If
End If
On Error GoTo 0

' Clase JsonParser simplificada para debug
Class JsonParser
    Private pos
    Private jsonText
    
    Public Function Parse(json)
        WScript.Echo "Iniciando Parse con: " & Left(json, 50) & "..."
        jsonText = json
        pos = 1
        
        ' Limpiar BOM y caracteres especiales al inicio
        Do While pos <= Len(jsonText)
            Dim firstChar
            firstChar = Mid(jsonText, pos, 1)
            Dim asciiVal
            asciiVal = Asc(firstChar)
            ' Saltar BOM UTF-8 y otros caracteres de control
            If asciiVal = 239 Or asciiVal = 187 Or asciiVal = 191 Or asciiVal < 32 Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        
        SkipWhitespace
        Set Parse = ParseValue()
    End Function
    
    Private Sub SkipWhitespace()
        Do While pos <= Len(jsonText)
            Dim char
            char = Mid(jsonText, pos, 1)
            If char = " " Or char = Chr(9) Or char = Chr(10) Or char = Chr(13) Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
    End Sub
    
    Private Function ParseValue()
        SkipWhitespace
        Dim char
        char = Mid(jsonText, pos, 1)
        WScript.Echo "ParseValue - carácter actual: '" & char & "' ASCII " & Asc(char)
        
        Select Case char
            Case "{"
                WScript.Echo "Detectado objeto - llamando ParseObject"
                Set ParseValue = ParseObject()
            Case Else
                Err.Raise 1001, "JsonParser", "Se requiere un objeto"
        End Select
    End Function
    
    Private Function ParseObject()
        WScript.Echo "Iniciando ParseObject"
        Dim obj
        Set obj = CreateObject("Scripting.Dictionary")
        pos = pos + 1 ' Saltar '{'
        SkipWhitespace
        
        If Mid(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Set ParseObject = obj
            WScript.Echo "Objeto vacío creado"
            Exit Function
        End If
        
        ' Parsear primer par clave-valor
        SkipWhitespace
        Dim key
        key = ParseString()
        WScript.Echo "Clave parseada: " & key
        
        Set ParseObject = obj
        WScript.Echo "Objeto creado con tipo: " & TypeName(obj)
    End Function
    
    Private Function ParseString()
        If Mid(jsonText, pos, 1) <> Chr(34) Then ' "
            Err.Raise 1003, "JsonParser", "Se esperaba una cadena"
        End If
        
        pos = pos + 1 ' Saltar primera comilla
        Dim startPos
        startPos = pos
        
        ' Buscar comilla de cierre
        Do While pos <= Len(jsonText)
            If Mid(jsonText, pos, 1) = Chr(34) Then ' "
                Exit Do
            End If
            pos = pos + 1
        Loop
        
        Dim result
        result = Mid(jsonText, startPos, pos - startPos)
        pos = pos + 1 ' Saltar comilla de cierre
        ParseString = result
    End Function
End Class