Attribute VB_Name = "modTestUtils"
Option Compare Database
Option Explicit

'******************************************************************************
' MÓDULO: modTestUtils
' DESCRIPCIÓN: Utilidades compartidas para el framework de pruebas.
'******************************************************************************



Public Sub VerifyAllTestTemplates()
    On Error GoTo errorHandler
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim projectPath As String
    projectPath = GetProjectPath()
    
    ' Lista de todas las plantillas requeridas por los tests de integración
    Dim templates As Variant
    templates = Array( _
        "back\test_db\templates\CONDOR_test_template.accdb", _
        "back\test_db\templates\Expedientes_test_template.accdb", _
        "back\test_db\templates\Lanzadera_test_template.accdb", _
        "back\test_db\templates\correos_test_template.accdb" _
    )
    
    Dim i As Integer
    Dim missingTemplates As String
    missingTemplates = ""
    
    For i = 0 To UBound(templates)
        Dim templatePath As String
        templatePath = projectPath & templates(i)
        
        If Not fs.FileExists(templatePath) Then
            If missingTemplates <> "" Then missingTemplates = missingTemplates & vbCrLf
            missingTemplates = missingTemplates & "- " & templatePath
        End If
    Next i
    
    If missingTemplates <> "" Then
        Err.Raise 53, "VerifyAllTestTemplates", "Las siguientes plantillas de base de datos no se encontraron:" & vbCrLf & missingTemplates & vbCrLf & vbCrLf & "Los tests de integración no pueden ejecutarse sin estas plantillas."
    End If
    
    Exit Sub
    
errorHandler:
    Err.Raise Err.Number, "modTestUtils.VerifyAllTestTemplates", Err.Description
End Sub

Public Sub PrepareTestDatabase(ByVal templatePath As String, ByVal activeTestPath As String)
    On Error GoTo errorHandler

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

errorHandler:
    ' Propagar el error para que el test que lo llamó falle con una descripción clara
    Err.Raise Err.Number, "modTestUtils.PrepareTestDatabase", Err.Description
End Sub

' ============================================================================
' SECCIÓN: GESTIÓN DEL CICLO DE VIDA DE LA SUITE DE PRUEBAS
' ============================================================================

Public Sub SuiteSetup(ByVal templateDbPath As String, ByVal activeDbPath As String)
    On Error GoTo ErrorHandler
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    If fs.FileExists(activeDbPath) Then
        fs.DeleteFile activeDbPath, True
    End If
    
    fs.CopyFile templateDbPath, activeDbPath
    
    Set fs = Nothing
    Exit Sub

ErrorHandler:
    Set fs = Nothing
    Err.Raise Err.Number, "modTestUtils.SuiteSetup", "No se pudo crear el entorno para la suite de pruebas: " & Err.Description
End Sub

Public Sub SuiteTeardown(ByVal activeDbPath As String)
    On Error GoTo ErrorHandler
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    If fs.FileExists(activeDbPath) Then
        fs.DeleteFile activeDbPath, True
    End If
    
    Set fs = Nothing
    Exit Sub

ErrorHandler:
    Set fs = Nothing
    Err.Raise Err.Number, "modTestUtils.SuiteTeardown", "No se pudo limpiar el entorno de la suite de pruebas: " & Err.Description
End Sub

Public Function GetProjectPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parentFolder As String
    parentFolder = fso.GetParentFolderName(CurrentProject.Path) ' Sube a /back
    parentFolder = fso.GetParentFolderName(parentFolder) ' Sube a / (raíz del proyecto)
    GetProjectPath = parentFolder & "\"
    Set fso = Nothing
End Function

