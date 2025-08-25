' Test de la función ExecuteSQLScript aislada
Dim objFSO, objArgs, strAccessPath
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
strAccessPath = "C:\test.accdb"

Sub ExecuteSQLScript()
    ' Declarar todas las variables al inicio
    Dim strSQLFilePath, objTextFile, strSQLContent
    Dim arrStatements, i, strStatement
    Dim db, dbEngine
    
    WScript.Echo "Función ExecuteSQLScript iniciada"
    
    ' Simular contenido SQL
    strSQLContent = "SELECT 1;SELECT 2;SELECT 3"
    
    ' Parsear el contenido: dividir por punto y coma
    arrStatements = Split(strSQLContent, ";")
    
    ' Ejecutar cada sentencia SQL (simulado)
    For i = 0 To UBound(arrStatements)
        strStatement = Trim(arrStatements(i))
        If Len(strStatement) > 0 Then
            WScript.Echo "Procesando: " & strStatement
        End If
    Next
    
    WScript.Echo "Función ExecuteSQLScript completada"
End Sub

' Llamar a la función
ExecuteSQLScript