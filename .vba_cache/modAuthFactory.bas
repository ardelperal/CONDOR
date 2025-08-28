Attribute VB_Name = "modAuthFactory"
Option Compare Database
Option Explicit


' =====================================================
' MODULO: modAuthFactory
' DESCRIPCION: Factory para la creación de servicios de autenticación
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

' Variable privada para almacenar el mock de AuthService
Private m_MockAuthService As IAuthService

' Función factory para crear y configurar el servicio de autenticación
Public Function CreateAuthService() As IAuthService
    On Error GoTo ErrorHandler
    
    ' Si hay un mock configurado, devolverlo
    If Not m_MockAuthService Is Nothing Then
        Set CreateAuthService = m_MockAuthService
        Exit Function
    End If
    
    ' Crear las dependencias necesarias usando sus respectivas factorías
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    Dim config As IConfig
    Set config = modConfig.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    Dim authRepository As IAuthRepository
    Set authRepository = modRepositoryFactory.CreateAuthRepository(config, errorHandler)
    
    ' Crear una instancia de la clase concreta
    Dim authServiceInstance As New CAuthService
    
    ' Inicializar la instancia concreta con todas las dependencias
    authServiceInstance.Initialize config, operationLogger, authRepository, errorHandler
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateAuthService = authServiceInstance
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modAuthFactory.CreateAuthService: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Método para inyectar un mock de AuthService (usado en pruebas unitarias)
Public Sub SetMockAuthService(mock As IAuthService)
    On Error GoTo ErrorHandler
    Set m_MockAuthService = mock
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modAuthFactory.SetMockAuthService: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Método para resetear el factory a su estado normal (usado en pruebas unitarias)
Public Sub ResetMock()
    On Error GoTo ErrorHandler
    Set m_MockAuthService = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print "Error en modAuthFactory.ResetMock: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


