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
Public Function CreateSolicitudService() As ISolicitudService
    On Error GoTo ErrorHandler
    
    ' 1. Crear dependencias de nivel más bajo primero, llamando a sus factorías SIN argumentos.
    Dim configService As IConfig
    Set configService = modConfigFactory.CreateConfigService()
    
    Dim errorHandlerService As IErrorHandlerService
    Set errorHandlerService = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim operationLoggerService As IOperationLogger
    Set operationLoggerService = modOperationLoggerFactory.CreateOperationLogger()
    
    ' 2. Crear el repositorio
    Dim solicitudRepo As ISolicitudRepository
    Set solicitudRepo = modRepositoryFactory.CreateSolicitudRepository()
    
    ' 3. Crear e inicializar la instancia del servicio
    Dim serviceInstance As New CSolicitudService
    serviceInstance.Initialize solicitudRepo, operationLoggerService, errorHandlerService
    
    ' 4. Devolver la instancia como el tipo de la interfaz
    Set CreateSolicitudService = serviceInstance
    
    Exit Function
    
ErrorHandler:
    ' Usar Debug.Print en una factoría es aceptable si errorHandler falla.
    Debug.Print "Error crítico en modSolicitudServiceFactory.CreateSolicitudService: " & Err.Description
    Set CreateSolicitudService = Nothing
End Function

