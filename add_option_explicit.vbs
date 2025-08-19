' Script para agregar Option Explicit a todos los módulos VBA
' Autor: CONDOR-Developer
' Fecha: Enero 2025

Dim fso, folder, file, content, lines, i, modified
Set fso = CreateObject("Scripting.FileSystemObject")

' Función para procesar un archivo
Sub ProcessFile(filePath)
    Dim content, lines, i, hasOptionExplicit, optionCompareLineIndex
    Dim newContent, line
    
    ' Leer el contenido del archivo
    Set file = fso.OpenTextFile(filePath, 1, False, -1) ' -1 = Unicode
    content = file.ReadAll
    file.Close
    
    ' Dividir en líneas
    lines = Split(content, vbCrLf)
    If UBound(lines) = 0 Then
        lines = Split(content, vbLf)
    End If
    
    hasOptionExplicit = False
    optionCompareLineIndex = -1
    
    ' Buscar Option Compare Database y Option Explicit
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If InStr(1, line, "Option Compare Database", vbTextCompare) > 0 Then
            optionCompareLineIndex = i
        End If
        If InStr(1, line, "Option Explicit", vbTextCompare) > 0 Then
            hasOptionExplicit = True
            Exit For
        End If
    Next
    
    ' Si tiene Option Compare Database pero no Option Explicit, agregarlo
    If optionCompareLineIndex >= 0 And Not hasOptionExplicit Then
        WScript.Echo "Agregando Option Explicit a: " & filePath
        
        ' Crear nuevo contenido
        newContent = ""
        For i = 0 To UBound(lines)
            newContent = newContent & lines(i)
            If i = optionCompareLineIndex Then
                newContent = newContent & vbCrLf & "Option Explicit"
            End If
            If i < UBound(lines) Then
                newContent = newContent & vbCrLf
            End If
        Next
        
        ' Escribir el archivo modificado
        Set file = fso.OpenTextFile(filePath, 2, True, -1) ' -1 = Unicode
        file.Write newContent
        file.Close
        
        modified = modified + 1
    End If
End Sub

' Función para procesar una carpeta recursivamente
Sub ProcessFolder(folderPath)
    Dim folder, file, subfolder
    Set folder = fso.GetFolder(folderPath)
    
    ' Procesar archivos en la carpeta actual
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "cls" Or _
           LCase(fso.GetExtensionName(file.Name)) = "bas" Then
            ProcessFile file.Path
        End If
    Next
    
    ' Procesar subcarpetas
    For Each subfolder In folder.SubFolders
        ProcessFolder subfolder.Path
    Next
End Sub

' Programa principal
modified = 0
WScript.Echo "Iniciando proceso de agregar Option Explicit..."
WScript.Echo "Procesando carpeta: C:\Proyectos\CONDOR\src"

ProcessFolder "C:\Proyectos\CONDOR\src"

WScript.Echo "Proceso completado. Archivos modificados: " & modified