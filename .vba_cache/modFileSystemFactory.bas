Attribute VB_Name = "modFileSystemFactory"
Option Compare Database
Option Explicit

' Módulo: modFileSystemFactory
' Descripción: Factory para crear instancias de IFileSystem
' Arquitectura: Capa de Servicios - Factory Pattern

' Crea una instancia de IFileSystem
' @param errorHandler: Servicio de manejo de errores (Opcional)
' @return IFileSystem: Instancia lista para usar
Public Function CreateFileSystem(Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IFileSystem
    On Error GoTo ErrorHandler
    
    Dim fileSystemInstance As New CFileSystem
    
    ' Devolver la interfaz
    Set CreateFileSystem = fileSystemInstance
    
    Exit Function
    
ErrorHandler:
    ' Usar el manejador de errores inyectado si está disponible, sino Debug.Print
    If Not errorHandler Is Nothing Then
        errorHandler.LogError Err.Number, Err.Description, "modFileSystemFactory.CreateFileSystem"
    Else
        Debug.Print "Error en modFileSystemFactory.CreateFileSystem: " & Err.Number & " - " & Err.Description
    End If
    Set CreateFileSystem = Nothing
End Function

