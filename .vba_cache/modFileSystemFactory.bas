Attribute VB_Name = "modFileSystemFactory"
Option Compare Database
Option Explicit


' Módulo: modFileSystemFactory
' Descripción: Factory para crear instancias de IFileSystem
' Arquitectura: Capa de Servicios - Factory Pattern

' Crea una instancia de IFileSystem
' @return IFileSystem: Instancia lista para usar
Public Function CreateFileSystem(Optional ByVal config As IConfig = Nothing) As IFileSystem
    On Error GoTo errorHandler
    
    Dim configService As IConfig
    If config Is Nothing Then
        ' Si no se provee una config (caso producción), la creamos.
        Set configService = modConfigFactory.CreateConfigService
    Else
        ' Si se provee (caso testing), la usamos.
        Set configService = config
    End If
    
    Dim fileSystemInstance As New CFileSystem
    ' Asumimos que CFileSystem tiene un método Initialize, siguiendo nuestro patrón estándar.
    ' Inyectamos la dependencia para que el objeto esté listo para usarse.
    fileSystemInstance.Initialize configService
    
    Set CreateFileSystem = fileSystemInstance
    
    Exit Function
    
errorHandler:
    Debug.Print "Error en modFileSystemFactory.CreateFileSystem: " & Err.Number & " - " & Err.Description
    Set CreateFileSystem = Nothing
End Function



