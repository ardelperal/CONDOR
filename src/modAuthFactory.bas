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
    ' Si hay un mock configurado, devolverlo
    If Not m_MockAuthService Is Nothing Then
        Set CreateAuthService = m_MockAuthService
        Exit Function
    End If
    
    ' Obtener la instancia de configuración usando el nuevo factory
    Dim config As IConfig: Set config = modConfig.CreateConfigService()
    
    ' Obtener la instancia del logger de operaciones
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger()
    
    ' Obtener la instancia del repositorio de solicitudes
    Dim solicitudRepository As ISolicitudRepository
    Set solicitudRepository = modRepositoryFactory.CreateSolicitudRepository()
    
    ' Crear una instancia de la clase concreta
    Dim authServiceInstance As New CAuthService
    
    ' Inicializar la instancia concreta con todas las dependencias
    authServiceInstance.Initialize config, operationLogger, solicitudRepository
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateAuthService = authServiceInstance
End Function

' Método para inyectar un mock de AuthService (usado en pruebas unitarias)
Public Sub SetMockAuthService(mock As IAuthService)
    Set m_MockAuthService = mock
End Sub

' Método para resetear el factory a su estado normal (usado en pruebas unitarias)
Public Sub ResetMock()
    Set m_MockAuthService = Nothing
End Sub
