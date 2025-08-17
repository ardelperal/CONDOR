Attribute VB_Name = "modTypes"
Option Compare Database
Option Explicit

' ============================================================================
' M?dulo: modTypes
' Descripci?n: Definici?n de tipos de datos personalizados para CONDOR
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Tipo que representa un expediente basado en la consulta SQL de integraci?n
' con la aplicaci?n de Expedientes existente
Public Type T_Expediente
    ID As Long                    ' Propiedad usada en modMockFramework
    IDExpediente As Long
    Nemotecnico As String
    Titulo As String
    ResponsableCalidad As String
    ResponsableTecnico As String  ' Se refiere al Jefe de Proyecto
    Pecal As String
End Type

' Tipo que representa una solicitud general
' Basado en la tabla Tb_Solicitudes de la especificaci?n funcional
Public Type T_Solicitud
    ID As Long                    ' Propiedad usada en modMockFramework
    NumeroExpediente As String
    TipoSolicitud As String
    EstadoInterno As String
    EstadoRAC As String
    FechaCreacion As Date
    FechaUltimaModificacion As Date
    Usuario As String
    Observaciones As String
    Activo As Boolean
End Type

' Tipo que representa los datos espec?ficos para Propuestas de Cambio (PC)
' Basado en la tabla TbDatos_PC de la especificaci?n funcional
Public Type T_Datos_PC
    ID As Long                    ' Propiedad usada en modMockFramework
    SolicitudID As Long           ' Propiedad usada en modMockFramework
    NumeroExpediente As String    ' Propiedad usada en modMockFramework
    TipoSolicitud As String       ' Propiedad usada en modMockFramework
    DescripcionCambio As String   ' Propiedad usada en modMockFramework
    JustificacionCambio As String ' Propiedad usada en modMockFramework
    ImpactoSeguridad As String    ' Propiedad usada en modMockFramework
    ImpactoCalidad As String      ' Propiedad usada en modMockFramework
    FechaCreacion As Date         ' Propiedad usada en modMockFramework
    FechaUltimaModificacion As Date
    Estado As String              ' Propiedad usada en modMockFramework
    Activo As Boolean             ' Propiedad usada en modMockFramework
End Type

' Tipo que representa los datos espec?ficos para Concesiones y Desviaciones (CD_CA)
' Basado en la tabla TbDatos_CD_CA de la especificaci?n funcional
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

' Tipo que representa los datos espec?ficos para Concesiones y Desviaciones de Sub-suministrador (CD_CA_SUB)
' Basado en la tabla TbDatos_CD_CA_SUB de la especificaci?n funcional
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



