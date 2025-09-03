Attribute VB_Name = "modSolicitudServiceFactory"
Option Compare Database
Option Explicit


'******************************************************************************
' MÓDULO: modSolicitudServiceFactory
' DESCRIPCIÓN: Factory para la inyección de dependencias del servicio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-14
'******************************************************************************

'******************************************************************************
' FACTORY METHODS
'******************************************************************************

'******************************************************************************
' FUNCIÓN: CreateSolicitudService
' DESCRIPCIÓN: Crea una instancia del servicio de solicitudes con todas sus dependencias
' RETORNA: ISolicitudService - Instancia del servicio completamente inicializada
'******************************************************************************
Public Function CreateSolicitudService(Optional ByVal config As IConfig = Nothing) As ISolicitudService
    On Error GoTo errorHandler
    
    ' 1. Determinar configuración final
    Dim finalConfig As IConfig
    If config Is Nothing Then
        Set finalConfig = modConfigFactory.CreateConfigService()
    Else
        Set finalConfig = config
    End If
    
    ' 2. Crear dependencias propagando la configuración
    Dim errorHandlerService As IErrorHandlerService
    Set errorHandlerService = modErrorHandlerFactory.CreateErrorHandlerService(finalConfig)
    
    Dim operationLoggerService As IOperationLogger
    Set operationLoggerService = modOperationLoggerFactory.CreateOperationLogger(finalConfig)
    
    ' 3. Crear el repositorio
    Dim solicitudRepo As ISolicitudRepository
    Set solicitudRepo = modRepositoryFactory.CreateSolicitudRepository(finalConfig)
    
    ' 4. Crear el servicio de autenticación
    Dim authService As IAuthService
    Set authService = modAuthFactory.CreateAuthService()
    
    ' 5. Crear e inicializar la instancia del servicio
    Dim serviceInstance As New CSolicitudService
    serviceInstance.Initialize solicitudRepo, operationLoggerService, errorHandlerService, authService
    
    ' 6. Devolver la instancia como el tipo de la interfaz
    Set CreateSolicitudService = serviceInstance
    
    Exit Function
    
errorHandler:
    ' Usar Debug.Print en una factoría es aceptable si errorHandler falla.
    Debug.Print "Error crítico en modSolicitudServiceFactory.CreateSolicitudService: " & Err.Description
    Set CreateSolicitudService = Nothing
End Function



