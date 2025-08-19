' Script para eliminar líneas Option Compare Database duplicadas
Dim objFSO, objFolder, objFile, objFiles
Dim strContent, arrLines, i, j
Dim bFoundFirst, strNewContent
Dim filesFixed

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Proyectos\CONDOR\src")
Set objFiles = objFolder.Files

filesFixed = 0

' Procesar archivos en src
For Each objFile In objFiles
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
        ProcessFile objFile.Path
    End If
Next

' Procesar archivos en src\mocks
Set objFolder = objFSO.GetFolder("C:\Proyectos\CONDOR\src\mocks")
Set objFiles = objFolder.Files
For Each objFile In objFiles
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
        ProcessFile objFile.Path
    End If
Next

' Procesar archivos en src\services
Set objFolder = objFSO.GetFolder("C:\Proyectos\CONDOR\src\services")
Set objFiles = objFolder.Files
For Each objFile In objFiles
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
        ProcessFile objFile.Path
    End If
Next

WScript.Echo "Proceso completado. Archivos corregidos: " & filesFixed

Sub ProcessFile(filePath)
    Dim objTextFile, strContent
    Dim arrLines, i
    Dim optionCompareFound, optionExplicitFound
    Dim hasChanges, newContent
    
    hasChanges = False
    optionCompareFound = False
    optionExplicitFound = False
    newContent = ""
    
    ' Leer todo el contenido del archivo
    Set objTextFile = objFSO.OpenTextFile(filePath, 1, False, 0) ' 0 = ANSI
    strContent = objTextFile.ReadAll
    objTextFile.Close
    
    ' Dividir en líneas
    arrLines = Split(strContent, vbCrLf)
    If UBound(arrLines) = 0 Then
        arrLines = Split(strContent, vbLf)
    End If
    
    ' Procesar cada línea
    For i = 0 To UBound(arrLines)
        If Trim(arrLines(i)) = "Option Compare Database" Then
            If Not optionCompareFound Then
                ' Primera vez que encontramos Option Compare Database
                optionCompareFound = True
                newContent = newContent & arrLines(i) & vbCrLf
            Else
                ' Es una línea duplicada, la omitimos
                hasChanges = True
                WScript.Echo "  [ELIMINANDO] Option Compare Database duplicada en " & objFSO.GetFileName(filePath)
            End If
        ElseIf Trim(arrLines(i)) = "Option Explicit" Then
            If Not optionExplicitFound Then
                ' Primera vez que encontramos Option Explicit
                optionExplicitFound = True
                newContent = newContent & arrLines(i) & vbCrLf
            Else
                ' Es una línea duplicada, la omitimos
                hasChanges = True
                WScript.Echo "  [ELIMINANDO] Option Explicit duplicada en " & objFSO.GetFileName(filePath)
            End If
        Else
            ' Mantener todas las demás líneas
            newContent = newContent & arrLines(i) & vbCrLf
        End If
    Next
    
    ' Si hubo cambios, reescribir el archivo
    If hasChanges Then
        ' Eliminar el último vbCrLf extra
        If Right(newContent, 2) = vbCrLf Then
            newContent = Left(newContent, Len(newContent) - 2)
        End If
        
        Set objTextFile = objFSO.OpenTextFile(filePath, 2, True, 0) ' 0 = ANSI
        objTextFile.Write newContent
        objTextFile.Close
        
        filesFixed = filesFixed + 1
        WScript.Echo "[CORREGIDO] " & objFSO.GetFileName(filePath)
    End If
End Sub