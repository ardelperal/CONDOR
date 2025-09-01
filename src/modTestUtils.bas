Attribute VB_Name = "modTestUtils"
Option Compare Database
Option Explicit
'******************************************************************************
' MÓDULO: modTestUtils
' DESCRIPCIÓN: Utilidades compartidas para el framework de pruebas.
'******************************************************************************

Public Function GetProjectPath() As String
    GetProjectPath = CurrentProject.Path & "\"
End Function

Public Sub PrepareTestDatabase(ByVal templatePath As String, ByVal activeTestPath As String)
    On Error GoTo ErrorHandler

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()

    ' Asegurarse de que la ruta completa a la plantilla existe
    If Not fs.FileExists(templatePath) Then
        Err.Raise 53, "PrepareTestDatabase", "El archivo plantilla de la base de datos no se encontró en: " & templatePath
    End If

    ' Borrar la base de datos de prueba activa si ya existe
    If fs.FileExists(activeTestPath) Then
        fs.DeleteFile activeTestPath, True ' True para forzar el borrado
    End If

    ' Crear el directorio de destino si no existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim destinationFolder As String
    destinationFolder = fso.GetParentFolderName(activeTestPath)
    Set fso = Nothing
    
    If Not fs.FolderExists(destinationFolder) Then
        fs.CreateFolder destinationFolder
    End If

    ' Copiar la plantilla para crear la nueva base de datos de prueba activa
    fs.CopyFile templatePath, activeTestPath

    Exit Sub

ErrorHandler:
    ' Propagar el error para que el test que lo llamó falle con una descripción clara
    Err.Raise Err.Number, "modTestUtils.PrepareTestDatabase", Err.Description
End Sub