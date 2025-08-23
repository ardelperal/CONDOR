Option Compare Database
Option Explicit
Implements IFileSystem

' =====================================================
' CLASE: CFileSystem
' DESCRIPCION: Implementación real del sistema de ficheros
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