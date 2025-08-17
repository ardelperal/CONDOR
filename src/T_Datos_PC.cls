VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Datos_PC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: T_Datos_PC
' DESCRIPCIÓN: Tipo de datos para almacenar información específica de Propuestas de Cambio (PC)
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas basándose en Test_SolicitudFactory.bas como especificación
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' PROPIEDADES PÚBLICAS
' Basadas en la función Test_CSolicitudPC_DatosPC_SetAndGet
'******************************************************************************

' Identificador único del registro de datos PC
Public ID As Long

' Referencia a la solicitud padre
Public IDSolicitud As Long

' Título de la Propuesta de Cambio
Public TituloPC As String

' Descripción detallada de la Propuesta de Cambio
Public DescripcionPC As String

' Justificación de la Propuesta de Cambio
Public JustificacionPC As String

' Análisis de impacto de la Propuesta de Cambio
Public ImpactoPC As String

' Fecha de creación del registro
Public FechaCreacion As Date

' Usuario que creó el registro
Public CreadoPor As String

'******************************************************************************
' PROPIEDADES ADICIONALES IDENTIFICADAS EN OTROS ARCHIVOS
' (Para compatibilidad con el sistema existente)
'******************************************************************************

' Propiedades técnicas específicas (identificadas en CMockSolicitudRepository)
Public Procesador As String
Public RAM As String
Public Almacenamiento As String
Public SistemaOperativo As String

'******************************************************************************
' PROPIEDADES IDENTIFICADAS EN modDatabase.bas
' (Utilizadas en la función SaveSolicitudPC)
'******************************************************************************

' ID de la solicitud padre (usado en SaveSolicitudPC)
Public SolicitudID As Long

' Número de expediente (copiado desde solicitud)
Public NumeroExpediente As String

' Tipo de solicitud (copiado desde solicitud)
Public TipoSolicitud As String

' Descripción del cambio propuesto
Public DescripcionCambio As String

' Justificación del cambio propuesto
Public JustificacionCambio As String

' Impacto en seguridad
Public ImpactoSeguridad As String

' Impacto en calidad
Public ImpactoCalidad As String

' Fecha de última modificación
Public FechaUltimaModificacion As Date

' Estado del registro de datos PC
Public Estado As String

' Indica si el registro está activo
Public Activo As Boolean