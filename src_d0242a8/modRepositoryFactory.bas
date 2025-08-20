Attribute VB_Name = "modRepositoryFactory"
'******************************************************************************
' M├ôDULO: modRepositoryFactory
' DESCRIPCI├ôN: Factory para la inyecci├│n de dependencias del repositorio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************

Option Explicit

'******************************************************************************
' FACTORY METHODS
'******************************************************************************

'******************************************************************************
' FUNCI├ôN: CreateSolicitudRepository
' DESCRIPCI├ôN: Crea una instancia del repositorio de solicitudes seg├║n el modo
' RETORNA: ISolicitudRepository - Instancia del repositorio (Mock o Real)
'******************************************************************************
Public Function CreateSolicitudRepository() As ISolicitudRepository
    ' Por ahora siempre devolvemos el Mock para las pruebas
    ' TODO: Implementar l├│gica para alternar entre Mock y Real seg├║n configuraci├│n
    Set CreateSolicitudRepository = New CMockSolicitudRepository
End Function
