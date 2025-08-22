Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modExpedienteServiceFactory
' Descripción: Factory para crear instancias de CExpedienteService con dependencias reales
' Autor: CONDOR-Developer
' Fecha: Enero 2025
' ============================================================================

' ============================================================================
' FUNCIÓN FACTORY PRINCIPAL
' ============================================================================

Public Function CreateExpedienteService() As CExpedienteService
    ' Crear instancia del servicio
    Dim expedienteService As New CExpedienteService
    
    ' Crear dependencias reales
    Dim config As IConfig
    Set config = New CConfig
    
    Dim logger As IOperationLogger
    Set logger = New COperationLogger
    
    Dim repository As ISolicitudRepository
    Set repository = New CSolicitudRepository
    
    ' Inyectar dependencias
    expedienteService.Initialize config, logger, repository
    
    ' Devolver instancia configurada
    Set CreateExpedienteService = expedienteService
End Function

' ============================================================================
' FUNCIÓN FACTORY PARA TESTING CON MOCKS
' ============================================================================

Public Function CreateExpedienteServiceWithMocks(ByVal mockConfig As IConfig, _
                                                ByVal mockLogger As IOperationLogger, _
                                                ByVal mockRepository As ISolicitudRepository) As CExpedienteService
    ' Crear instancia del servicio
    Dim expedienteService As New CExpedienteService
    
    ' Inyectar mocks
    expedienteService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Devolver instancia configurada con mocks
    Set CreateExpedienteServiceWithMocks = expedienteService
End Function