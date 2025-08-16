Attribute VB_Name = "modRepositoryFactory"
'******************************************************************************
' MÓDULO: modRepositoryFactory
' DESCRIPCIÓN: Factory para la inyección de dependencias del repositorio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************

Option Explicit

'******************************************************************************
' FACTORY METHODS
'******************************************************************************

'******************************************************************************
' FUNCIÓN: CreateSolicitudRepository
' DESCRIPCIÓN: Crea una instancia del repositorio de solicitudes según el modo
' RETORNA: ISolicitudRepository - Instancia del repositorio (Mock o Real)
'******************************************************************************
Public Function CreateSolicitudRepository() As ISolicitudRepository
    ' Por ahora siempre devolvemos el Mock para las pruebas
    ' TODO: Implementar lógica para alternar entre Mock y Real según configuración
    Set CreateSolicitudRepository = New CMockSolicitudRepository
End Function