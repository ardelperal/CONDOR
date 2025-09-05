Attribute VB_Name = "modFileSystemFactory"
Option Compare Database
Option Explicit


' Módulo: modFileSystemFactory
' Descripción: Factory para crear instancias de IFileSystem
' Arquitectura: Capa de Servicios - Factory Pattern

' Crea una instancia de IFileSystem
' @return IFileSystem: Instancia lista para usar
Public Function CreateFileSystem() As IFileSystem
    On Error GoTo errorHandler
    
    Dim config As IConfig
    Set config = modTestContext.GetTestConfig()
    
    Dim fileSystemInstance As New CFileSystem
    ' Asumimos que CFileSystem tiene un método Initialize, siguiendo nuestro patrón estándar.
    ' Inyectamos la dependencia para que el objeto esté listo para usarse.
    fileSystemInstance.Initialize config
    
    Set CreateFileSystem = fileSystemInstance
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modFileSystemFactory.CreateFileSystem: " & Err.Number & " - " & Err.Description
    Set CreateFileSystem = Nothing
End Function



