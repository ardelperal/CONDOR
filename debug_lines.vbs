Dim objFSO, objFile, strLine, lineNum
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Proyectos\CONDOR\src\T_Datos_CD_CA_SUB.cls", 1, False, 0)

lineNum = 1
Do While Not objFile.AtEndOfStream And lineNum <= 15
    strLine = objFile.ReadLine
    WScript.Echo lineNum & ": '" & strLine & "' (len=" & Len(strLine) & ")"
    lineNum = lineNum + 1
Loop

objFile.Close