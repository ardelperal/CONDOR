Attribute VB_Name = "modFileSystemFactory"
Option Compare Database
Option Explicit
' Módulo: modFileSystemFactory
' Descripción: Factory para crear instancias de IFileSystem
' Arquitectura: Capa de Servicios - Factory Pattern

' Crea una instancia de IFileSystem
' @return IFileSystem: Instancia lista para usar
Public Function CreateFileSystem() As IFileSystem
    On Error GoTo ErrorHandler
    
    Dim fileSystemInstance As New CFileSystem
    
    ' Devolver la interfaz
    Set CreateFileSystem = fileSystemInstance
    
    Exit Function
    
ErrorHandler:
    ' En caso de error en el factory, usamos logging directo para evitar recursión
    Debug.Print "Error en modFileSystemFactory.CreateFileSystem: " & Err.Number & " - " & Err.Description
    Set CreateFileSystem = Nothing
End Function