Option Compare Database
Option Explicit

' Clase: CMockTextStream
' Descripción: Mock del objeto TextStream para pruebas unitarias aisladas
' Arquitectura: Capa de Testing - Mock
' Propósito: Simular operaciones de escritura de ficheros sin interactuar con ficheros reales

Public ParentMock As CMockFileSystemObject

' Mock del método WriteLine
Public Sub WriteLine(ByVal content As String)
    On Error GoTo ErrorHandler
    
    If Not ParentMock Is Nothing Then
        ParentMock.CaptureWrittenContent content
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CMockTextStream.WriteLine: " & Err.Description
End Sub

' Mock del método Close
Public Sub Close()
    On Error GoTo ErrorHandler
    
    If Not ParentMock Is Nothing Then
        ParentMock.CaptureClose
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CMockTextStream.Close: " & Err.Description
End Sub

Private Sub Class_Terminate()
    Set ParentMock = Nothing
End Sub