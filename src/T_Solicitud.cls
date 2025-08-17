VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: T_Solicitud
' DESCRIPCIÓN: Tipo de datos para almacenar información de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas basándose en modDatabase.bas como fuente de verdad
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' PROPIEDADES PÚBLICAS
' Basadas en el análisis de modDatabase.bas función SaveSolicitudPC
'******************************************************************************

' Identificador único de la solicitud
Public ID As Long

' Número de expediente asociado
Public NumeroExpediente As String

' Tipo de solicitud (PC, CD/CA, etc.)
Public TipoSolicitud As String

' Estado interno de la solicitud
Public EstadoInterno As String

' Estado en el sistema RAC
Public EstadoRAC As String

' Fecha de creación de la solicitud
Public FechaCreacion As Date

' Fecha de última modificación
Public FechaUltimaModificacion As Date

' Usuario que maneja la solicitud
Public Usuario As String

' Observaciones de la solicitud
Public Observaciones As String

' Indica si la solicitud está activa
Public Activo As Boolean

'******************************************************************************
' PROPIEDADES ADICIONALES IDENTIFICADAS EN GetSolicitudData
'******************************************************************************

' Código de la solicitud (identificado en CSolicitudPC)
Public CodigoSolicitud As String

' Referencia al expediente (identificado en CSolicitudPC como IDExpediente)
Public IDExpediente As String

' PROPIEDAD CRÍTICA IDENTIFICADA EN TESTS
' Identificador de solicitud usado en CSolicitudPC y tests
Public idSolicitud As Long