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
    idExpediente As Long
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
    tipoSolicitud As String
    estadoInterno As String
    fechaCreacion As Date
    FechaUltimaModificacion As Date
    usuario As String
    Observaciones As String
    Activo As Boolean
End Type

' NOTA: Los tipos T_Datos_PC, T_Datos_CD_CA y T_Datos_CD_CA_SUB han sido
' reemplazados por clases (.cls) que implementan la especificaci?n completa
' del CONDOR_MASTER_PLAN.md. Las clases proporcionan mejor encapsulaci?n
' y siguen las especificaciones de las tablas de base de datos.










