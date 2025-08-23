Option Compare Database
Option Explicit

#If DEV_MODE Then

' Clase: CMockFileSystemObject
' Descripción: Mock del FileSystemObject para pruebas unitarias aisladas
' Arquitectura: Capa de Testing - Mock
' Propósito: Simular operaciones de sistema de ficheros sin interactuar con ficheros reales

' Propiedades públicas para capturar las operaciones
Public LastWrittenContent As String
Public LastFilePath As String
Public OpenTextFileCallCount As Long
Public WriteLineCallCount As Long
Public CloseCallCount As Long

' Mock del objeto TextStream
Private m_mockTextStream As CMockTextStream

Private Sub Class_Initialize()
    Reset
End Sub

Public Sub Reset()
    LastWrittenContent = ""
    LastFilePath = ""
    OpenTextFileCallCount = 0
    WriteLineCallCount = 0
    CloseCallCount = 0
    Set m_mockTextStream = New CMockTextStream
End Sub

' Mock del método OpenTextFile
Public Function OpenTextFile(ByVal fileName As String, Optional ByVal ioMode As Long = 1, Optional ByVal create As Boolean = False) As Object
    On Error GoTo ErrorHandler
    
    OpenTextFileCallCount = OpenTextFileCallCount + 1
    LastFilePath = fileName
    
    ' Configurar el mock TextStream para que capture el contenido
    Set m_mockTextStream.ParentMock = Me
    Set OpenTextFile = m_mockTextStream
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en CMockFileSystemObject.OpenTextFile: " & Err.Description
End Function

' Método interno para capturar el contenido escrito
Friend Sub CaptureWrittenContent(ByVal content As String)
    If LastWrittenContent = "" Then
        LastWrittenContent = content
    Else
        LastWrittenContent = LastWrittenContent & vbCrLf & content
    End If
    WriteLineCallCount = WriteLineCallCount + 1
End Sub

' Método interno para capturar el cierre del fichero
Friend Sub CaptureClose()
    CloseCallCount = CloseCallCount + 1
End Sub

' Métodos de verificación para las pruebas
Public Function WasOpenTextFileCalled() As Boolean
    WasOpenTextFileCalled = (OpenTextFileCallCount > 0)
End Function

Public Function WasWriteLineCalled() As Boolean
    WasWriteLineCalled = (WriteLineCallCount > 0)
End Function

Public Function WasCloseCalled() As Boolean
    WasCloseCalled = (CloseCallCount > 0)
End Function

Public Function GetLastWrittenContent() As String
    GetLastWrittenContent = LastWrittenContent
End Function

Public Function GetLastFilePath() As String
    GetLastFilePath = LastFilePath
End Function

#End If