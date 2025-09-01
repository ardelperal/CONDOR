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

' =====================================================
' CLASE: CMockFileSystem
' DESCRIPCION: Mock del sistema de ficheros para pruebas unitarias
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 1.0 - Mock para aislamiento de pruebas
' =====================================================

' Mock del sistema de ficheros que permite verificar
' las llamadas realizadas durante las pruebas unitarias
' sin interactuar con el sistema de ficheros real

' Propiedades públicas para verificación en pruebas
Public WasOpenTextFileCalled As Boolean
Public LastPath As String
Public LastMode As Long
Public LastCreate As Boolean
Public CallCount As Long

' --- Propiedades públicas para verificación en pruebas ---
Public FileExists_ReturnValue As Boolean
Public FolderExists_ReturnValue As Boolean
Public FileExists_WasCalled As Boolean
Public FolderExists_WasCalled As Boolean
Public CopyFile_WasCalled As Boolean
Public DeleteFile_WasCalled As Boolean
Public CreateFolder_WasCalled As Boolean
Public DeleteFolder_WasCalled As Boolean

' --- Propiedades específicas para DeleteFile ---
Public LastDeletedPath As String
Public DeleteWasForced As Boolean
Public DeleteFileCallCount As Long

' Mock del objeto TextFile que se devuelve
Private m_mockTextFile As CMockTextFile

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Set m_mockTextFile = New CMockTextFile
    Reset
End Sub

Private Sub Class_Terminate()
    Set m_mockTextFile = Nothing
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS PARA CONFIGURACIÓN DEL MOCK
' ============================================================================

' Reinicia el estado del mock para una nueva prueba
Public Sub Reset()
    WasOpenTextFileCalled = False
    LastPath = ""
    LastMode = 0
    LastCreate = False
    CallCount = 0
    m_mockTextFile.Reset
    
    ' Reiniciar nuevas propiedades
    FileExists_ReturnValue = False
    FolderExists_ReturnValue = False
    FileExists_WasCalled = False
    FolderExists_WasCalled = False
    CopyFile_WasCalled = False
    DeleteFile_WasCalled = False
    CreateFolder_WasCalled = False
    DeleteFolder_WasCalled = False
    
    ' Reiniciar propiedades específicas de DeleteFile
    LastDeletedPath = ""
    DeleteWasForced = False
    DeleteFileCallCount = 0
End Sub

' Obtiene el mock del TextFile para configuración adicional
Public Function GetMockTextFile() As CMockTextFile
    Set GetMockTextFile = m_mockTextFile
End Function

' ============================================================================
' IMPLEMENTACIÃ“N DE LA INTERFAZ IFileSystem
' ============================================================================

' Mock de la apertura de archivos de texto
Private Function IFileSystem_OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    ' Registrar la llamada para verificación
    WasOpenTextFileCalled = True
    LastPath = Path
    LastMode = Mode
    LastCreate = Create
    CallCount = CallCount + 1
    
    ' Devolver el mock del TextFile
    Set IFileSystem_OpenTextFile = m_mockTextFile
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' --- Implementación de la interfaz ---

Private Function IFileSystem_FileExists(ByVal Path As String) As Boolean
    FileExists_WasCalled = True
    IFileSystem_FileExists = FileExists_ReturnValue
End Function

Private Function IFileSystem_FolderExists(ByVal Path As String) As Boolean
    FolderExists_WasCalled = True
    IFileSystem_FolderExists = FolderExists_ReturnValue
End Function

Private Sub IFileSystem_CopyFile(ByVal Source As String, ByVal Destination As String)
    CopyFile_WasCalled = True
End Sub

Private Sub IFileSystem_DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False)
    DeleteFile_WasCalled = True
    LastDeletedPath = Path
    DeleteWasForced = Force
    DeleteFileCallCount = DeleteFileCallCount + 1
End Sub

Private Sub IFileSystem_CreateFolder(ByVal Path As String)
    CreateFolder_WasCalled = True
End Sub

Private Sub IFileSystem_DeleteFolder(ByVal Path As String)
    DeleteFolder_WasCalled = True
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para OpenTextFile
Public Function OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    Set OpenTextFile = IFileSystem_OpenTextFile(Path, Mode, Create)
End Function

' Método público de conveniencia para FileExists
Public Function FileExists(ByVal Path As String) As Boolean
    FileExists = IFileSystem_FileExists(Path)
End Function

' Método público de conveniencia para FolderExists
Public Function FolderExists(ByVal Path As String) As Boolean
    FolderExists = IFileSystem_FolderExists(Path)
End Function

' Método público de conveniencia para CopyFile
Public Sub CopyFile(ByVal Source As String, ByVal Destination As String)
    IFileSystem_CopyFile Source, Destination
End Sub

' Método público de conveniencia para DeleteFile
Public Sub DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False)
    IFileSystem_DeleteFile Path, Force
End Sub

' Método público de conveniencia para CreateFolder
Public Sub CreateFolder(ByVal Path As String)
    IFileSystem_CreateFolder Path
End Sub

' Método público de conveniencia para DeleteFolder
Public Sub DeleteFolder(ByVal Path As String)
    IFileSystem_DeleteFolder Path
End Sub


