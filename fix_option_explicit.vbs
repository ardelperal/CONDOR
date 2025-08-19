' Script para agregar Option Explicit a todos los archivos VBA
Dim fso, folder, files, file
Dim srcPath

srcPath = "c:\Proyectos\CONDOR\src"
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(srcPath)

For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "cls" Or LCase(fso.GetExtensionName(file.Name)) = "bas" Then
        ProcessFile file.Path
    End If
Next

WScript.Echo "Procesamiento completado."

Sub ProcessFile(filePath)
    Dim fso, file, content, lines, i, modified
    Dim hasOptionCompare, hasOptionExplicit, insertIndex
    Dim newLines()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(filePath, 1) ' Leer como texto plano
    content = file.ReadAll
    file.Close
    
    ' Dividir en líneas
    lines = Split(content, vbCrLf)
    If UBound(lines) = 0 Then
        lines = Split(content, vbLf)
    End If
    
    hasOptionCompare = False
    hasOptionExplicit = False
    insertIndex = -1
    modified = False
    
    ' Buscar Option Compare Database y Option Explicit
    For i = 0 To UBound(lines)
        Dim lineText
        lineText = Trim(lines(i))
        
        If lineText = "Option Compare Database" Then
            hasOptionCompare = True
            insertIndex = i
        End If
        If lineText = "Option Explicit" Then
            hasOptionExplicit = True
        End If
    Next
    
    ' Si tiene Option Compare Database pero no Option Explicit, agregarlo
    If hasOptionCompare And Not hasOptionExplicit Then
        WScript.Echo "Agregando Option Explicit a: " & fso.GetFileName(filePath)
        
        ' Crear nuevo array con Option Explicit insertado
        ReDim newLines(UBound(lines) + 1)
        
        ' Copiar líneas hasta Option Compare Database
        For i = 0 To insertIndex
            newLines(i) = lines(i)
        Next
        
        ' Insertar Option Explicit
        newLines(insertIndex + 1) = "Option Explicit"
        
        ' Copiar el resto de líneas
        For i = insertIndex + 1 To UBound(lines)
            newLines(i + 1) = lines(i)
        Next
        
        ' Escribir archivo modificado
        content = Join(newLines, vbCrLf)
        Set file = fso.OpenTextFile(filePath, 2, True) ' Escribir como texto plano
        file.Write content
        file.Close
        
        modified = True
    End If
    
    If Not modified Then
        WScript.Echo "Sin cambios en: " & fso.GetFileName(filePath)
    End If
End Sub