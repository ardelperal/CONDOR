VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Implements IFileSystem

' =====================================================
' CLASE: CFileSystem
' DESCRIPCIÓN: Implementación real del sistema de ficheros
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 1.0 - Wrapper del FileSystemObject
' =====================================================

' Wrapper alrededor del FileSystemObject real
' Permite abstraer las operaciones de sistema de ficheros
' para facilitar las pruebas unitarias con mocks

Private fso As Object ' Scripting.FileSystemObject
Private m_config As IConfig

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Terminate()
    Set fso = Nothing
    Set m_config = Nothing
End Sub

' ============================================================================
' INICIALIZACIÓN CON INYECCIÓN DE DEPENDENCIAS
' ============================================================================

Public Sub Initialize(ByVal config As IConfig)
    Set m_config = config
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IFileSystem
' ============================================================================

' Implementa la apertura de archivos de texto
Private Function IFileSystem_OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    Set IFileSystem_OpenTextFile = fso.OpenTextFile(Path, Mode, Create)
End Function

' Implementa la verificación de existencia de archivos
Private Function IFileSystem_FileExists(ByVal Path As String) As Boolean
    IFileSystem_FileExists = fso.FileExists(Path)
End Function

' Implementa la verificación de existencia de carpetas
Private Function IFileSystem_FolderExists(ByVal Path As String) As Boolean
    IFileSystem_FolderExists = fso.FolderExists(Path)
End Function

' Implementa la copia de archivos
Private Sub IFileSystem_CopyFile(ByVal Source As String, ByVal Destination As String)
    fso.CopyFile Source, Destination
End Sub

' Implementa la eliminación de archivos
Private Sub IFileSystem_DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False)
    fso.DeleteFile Path, Force
End Sub

' Implementa la creación de carpetas
Private Sub IFileSystem_CreateFolder(ByVal Path As String)
    fso.CreateFolder Path
End Sub

' Implementa la eliminación de carpetas
Private Sub IFileSystem_DeleteFolder(ByVal Path As String)
    fso.DeleteFolder Path
End Sub

' Implementa la eliminación recursiva de carpetas
Private Sub IFileSystem_DeleteFolderRecursive(ByVal Path As String)
    On Error Resume Next
    If fso.FolderExists(Path) Then
        fso.DeleteFolder Path, True ' El True fuerza el borrado recursivo
    End If
    On Error GoTo 0
End Sub

' Implementa la escritura de líneas a archivo
Private Sub IFileSystem_WriteLineToFile(ByVal filePath As String, ByVal line As String)
    ' Encapsula la lógica de abrir, escribir y cerrar.
    Dim file As Object
    ' 8 = ForAppending, True = Create if not exists
    Set file = fso.OpenTextFile(filePath, 8, True)
    file.WriteLine line
    file.Close
    Set file = Nothing
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Function OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    Set OpenTextFile = IFileSystem_OpenTextFile(Path, Mode, Create)
End Function

Public Function FileExists(ByVal Path As String) As Boolean
    FileExists = IFileSystem_FileExists(Path)
End Function

Public Function FolderExists(ByVal Path As String) As Boolean
    FolderExists = IFileSystem_FolderExists(Path)
End Function

Public Sub CopyFile(ByVal Source As String, ByVal Destination As String)
    IFileSystem_CopyFile Source, Destination
End Sub

Public Sub DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False)
    IFileSystem_DeleteFile Path, Force
End Sub

Public Sub CreateFolder(ByVal Path As String)
    IFileSystem_CreateFolder Path
End Sub

Public Sub DeleteFolder(ByVal Path As String)
    IFileSystem_DeleteFolder Path
End Sub

Public Sub DeleteFolderRecursive(ByVal Path As String)
    ' Delega la llamada a la implementación privada y segura de la interfaz
    Call IFileSystem_DeleteFolderRecursive(Path)
End Sub

Public Sub WriteLineToFile(ByVal filePath As String, ByVal line As String)
    Call IFileSystem_WriteLineToFile(filePath, line)
End Sub




