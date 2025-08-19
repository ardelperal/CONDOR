' Test script for CleanVBAFile function
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Include the CleanVBAFile function (simplified version)
Function TestCleanVBAFile(filePath)
    Dim objStream, strContent, arrLines, i, cleanedLines
    Dim strLine
    
    ' Read file content
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
    
    ' Normalize line breaks
    strContent = Replace(strContent, vbCrLf, vbLf)
    strContent = Replace(strContent, vbCr, vbLf)
    arrLines = Split(strContent, vbLf)
    
    ' Process lines
    cleanedLines = ""
    
    For i = 0 To UBound(arrLines)
        strLine = Trim(arrLines(i))
        
        ' Check if line should be removed
        Dim shouldRemove
        shouldRemove = False
        
        If Left(strLine, 7) = "VERSION" Then shouldRemove = True
        If Left(strLine, 5) = "BEGIN" Then shouldRemove = True
        If Left(strLine, 8) = "MultiUse" Then shouldRemove = True
        If strLine = "END" Then shouldRemove = True
        If Left(strLine, 9) = "Attribute" Then shouldRemove = True
        ' IMPORTANTE: Solo eliminar "Option " (con espacio) para no afectar "Optional"
        If Left(strLine, 7) = "Option " Then shouldRemove = True
        If strLine = "" And cleanedLines = "" Then shouldRemove = True
        
        If Not shouldRemove Then
            strLine = arrLines(i) ' Use original line with spaces
            If cleanedLines <> "" Then
                cleanedLines = cleanedLines & vbCrLf
            End If
            cleanedLines = cleanedLines & strLine
        End If
    Next
    
    TestCleanVBAFile = cleanedLines
End Function

' Test the function
Dim result
result = TestCleanVBAFile("C:\Proyectos\CONDOR\test_optional_param.cls")
WScript.Echo "=== CLEANED CONTENT ==="
WScript.Echo result
WScript.Echo "=== END CLEANED CONTENT ==="