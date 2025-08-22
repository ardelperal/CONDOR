Attribute VB_Name = "modAuthFactory"
Option Compare Database
Option Explicit

' =====================================================
' MODULO: modAuthFactory
' DESCRIPCION: Factory para la creación de servicios de autenticación
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

' Función factory para crear y configurar el servicio de autenticación
Public Function CreateAuthService() As IAuthService
    ' Obtener la instancia de configuración
    Dim config As CConfig: Set config = modConfig.GetInstance()
    
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