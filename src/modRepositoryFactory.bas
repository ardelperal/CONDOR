Attribute VB_Name = "modRepositoryFactory"
'******************************************************************************
' MÓDULO: modRepositoryFactory
' DESCRIPCIÓN: Factory para la inyección de dependencias del repositorio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************

#If DEV_MODE Then

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
    #If DEV_MODE Then
        ' En modo desarrollo, devolvemos el Mock para las pruebas
        Set CreateSolicitudRepository = New CMockSolicitudRepository
    #Else
        ' En modo producción, devolvemos la implementación real
        Set CreateSolicitudRepository = New CSolicitudRepository
    #End If
End Function

#End If