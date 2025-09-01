VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IFileSystem

Public WasOpenTextFileCalled As Boolean
Public LastPath As String
Private m_mockTextFile As CMockTextFile

Private Sub Class_Initialize()
    Set m_mockTextFile = New CMockTextFile
    Reset
End Sub

Public Sub Reset()
    WasOpenTextFileCalled = False
    LastPath = ""
    m_mockTextFile.Reset
End Sub

Public Function GetMockTextFile() As CMockTextFile
    Set GetMockTextFile = m_mockTextFile
End Function

Private Sub IFileSystem_CreateFolder(ByVal Path As String): End Sub

Private Sub IFileSystem_CopyFile(ByVal Source As String, ByVal Destination As String): End Sub

Private Sub IFileSystem_DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False): End Sub

Private Sub IFileSystem_DeleteFolder(ByVal Path As String): End Sub

Private Sub IFileSystem_DeleteFolderRecursive(ByVal Path As String): End Sub

Private Function IFileSystem_FileExists(ByVal Path As String) As Boolean: End Function

Private Function IFileSystem_FolderExists(ByVal Path As String) As Boolean: End Function

Private Function IFileSystem_OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    WasOpenTextFileCalled = True
    LastPath = Path
    Set IFileSystem_OpenTextFile = m_mockTextFile
End Function




