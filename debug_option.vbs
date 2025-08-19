' Script de depuración simple para verificar Option Explicit
Dim fso, file, content, lines, i
Dim hasOptionCompare, hasOptionExplicit

Set fso = CreateObject("Scripting.FileSystemObject")

' Leer archivo como texto plano
Set file = fso.OpenTextFile("c:\Proyectos\CONDOR\src\Test_Codificacion.bas", 1)
content = file.ReadAll
file.Close

lines = Split(content, vbCrLf)
If UBound(lines) = 0 Then
    lines = Split(content, vbLf)
End If

hasOptionCompare = False
hasOptionExplicit = False

WScript.Echo "Analizando Test_Codificacion.bas:"
WScript.Echo "Total de líneas: " & (UBound(lines) + 1)
WScript.Echo ""

' Mostrar las primeras 10 líneas
For i = 0 To 9
    If i <= UBound(lines) Then
        Dim lineText
        lineText = Trim(lines(i))
        WScript.Echo "Línea " & (i + 1) & ": [" & lineText & "]"
        
        If lineText = "Option Compare Database" Then
            hasOptionCompare = True
            WScript.Echo "  -> Encontrado Option Compare Database"
        End If
        If lineText = "Option Explicit" Then
            hasOptionExplicit = True
            WScript.Echo "  -> Encontrado Option Explicit"
        End If
    End If
Next

WScript.Echo ""
WScript.Echo "Resultado:"
WScript.Echo "hasOptionCompare: " & hasOptionCompare
WScript.Echo "hasOptionExplicit: " & hasOptionExplicit
WScript.Echo "Necesita Option Explicit: " & (hasOptionCompare And Not hasOptionExplicit)