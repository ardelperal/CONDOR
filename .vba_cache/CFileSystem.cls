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

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
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


