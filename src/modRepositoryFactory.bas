Attribute VB_Name = "modRepositoryFactory"
Option Compare Database
Option Explicit

' =====================================================
' FACTORY: modRepositoryFactory
' DESCRIPCIÓN: Crea instancias de repositorios.
' PATRÓN ARQUITECTÓNICO: Factory con inyección de dependencias opcional para testing.
' =====================================================

' Flag para alternar entre implementaciones reales y mocks
Public Const DEV_MODE As Boolean = True

Public Function CreateExpedienteRepository(Optional ByVal config As IConfig = Nothing, Optional ByVal errorHandler As IErrorHandlerService = Nothing) As IExpedienteRepository
    On Error GoTo ErrorHandler
    
    Dim repoImpl As New CExpedienteRepository
    
    ' CORRECCIÓN ARQUITECTÓNICA: Usar dependencias inyectadas si existen (para tests),
    ' o crearlas si no (para la aplicación).
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        Set effectiveConfig = modConfigFactory.CreateConfigService()
    Else
        Set effectiveConfig = config
    End If
    
    Dim effectiveErrorHandler As IErrorHandlerService
    If errorHandler Is Nothing Then
        Set effectiveErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Else
        Set effectiveErrorHandler = errorHandler
    End If
    
    repoImpl.Initialize effectiveConfig, effectiveErrorHandler
    
    Set CreateExpedienteRepository = repoImpl
    
    Exit Function
ErrorHandler:
    ' Implementar manejo de errores robusto
    Debug.Print "Error en modRepositoryFactory.CreateExpedienteRepository: " & Err.Description
    Set CreateExpedienteRepository = Nothing
End Function

' Añadir aquí el resto de funciones Create... para otros repositorios (ISolicitudRepository, etc.)
' siguiendo el mismo patrón de parámetros opcionales.