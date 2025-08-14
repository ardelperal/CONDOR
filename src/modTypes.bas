Attribute VB_Name = "modTypes"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modTypes
' Descripción: Definición de tipos de datos personalizados para CONDOR
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Tipo que representa un expediente basado en la consulta SQL de integración
' con la aplicación de Expedientes existente
Public Type T_Expediente
    IDExpediente As Long
    Nemotecnico As String
    Titulo As String
    ResponsableCalidad As String
    ResponsableTecnico As String  ' Se refiere al Jefe de Proyecto
    Pecal As String
End Type

' Tipo que representa los datos específicos para Propuestas de Cambio (PC)
' Basado en la tabla TbDatos_PC de la especificación funcional
Public Type T_Datos_PC
    ID As Long
    SolicitudID As Long
    NumeroExpediente As String
    TipoSolicitud As String
    DescripcionCambio As String
    JustificacionCambio As String
    ImpactoSeguridad As String
    ImpactoCalidad As String
    FechaCreacion As Date
    FechaUltimaModificacion As Date
    Estado As String
    Activo As Boolean
End Type

' Tipo que representa los datos específicos para Concesiones y Desviaciones (CD_CA)
' Basado en la tabla TbDatos_CD_CA de la especificación funcional
Public Type T_Datos_CD_CA
    ID As Long
    SolicitudID As Long
    NumeroExpediente As String
    TipoSolicitud As String
    DescripcionDesviacion As String
    JustificacionDesviacion As String
    ImpactoSeguridad As String
    ImpactoCalidad As String
    MedidasCorrectivas As String
    FechaCreacion As Date
    FechaUltimaModificacion As Date
    Estado As String
    Activo As Boolean
End Type

' Tipo que representa los datos específicos para Concesiones y Desviaciones de Sub-suministrador (CD_CA_SUB)
' Basado en la tabla TbDatos_CD_CA_SUB de la especificación funcional
Public Type T_Datos_CD_CA_SUB
    ID As Long
    SolicitudID As Long
    NumeroExpediente As String
    TipoSolicitud As String
    NombreSubsuministrador As String
    DescripcionDesviacion As String
    JustificacionDesviacion As String
    ImpactoSeguridad As String
    ImpactoCalidad As String
    MedidasCorrectivas As String
    FechaCreacion As Date
    FechaUltimaModificacion As Date
    Estado As String
    Activo As Boolean
End Type