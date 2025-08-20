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
' DESCRIPCI├ôN: Tipo de datos para almacenar informaci├│n espec├¡fica de Propuestas de Cambio (PC)
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas bas├índose en Test_SolicitudFactory.bas como especificaci├│n
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' PROPIEDADES P├ÜBLICAS
' Basadas en la funci├│n Test_CSolicitudPC_DatosPC_SetAndGet
'******************************************************************************

' Identificador ├║nico del registro de datos PC
Public ID As Long

' Referencia a la solicitud padre
Public IDSolicitud As Long

' T├¡tulo de la Propuesta de Cambio
Public TituloPC As String

' Descripci├│n detallada de la Propuesta de Cambio
Public DescripcionPC As String

' Justificaci├│n de la Propuesta de Cambio
Public JustificacionPC As String

' An├ílisis de impacto de la Propuesta de Cambio
Public ImpactoPC As String

' Fecha de creaci├│n del registro
Public FechaCreacion As Date

' Usuario que cre├│ el registro
Public CreadoPor As String

'******************************************************************************
' PROPIEDADES ADICIONALES IDENTIFICADAS EN OTROS ARCHIVOS
' (Para compatibilidad con el sistema existente)
'******************************************************************************

' Propiedades t├®cnicas espec├¡ficas (identificadas en CMockSolicitudRepository)
Public Procesador As String
Public RAM As String
Public Almacenamiento As String
Public SistemaOperativo As String

'******************************************************************************
' PROPIEDADES IDENTIFICADAS EN modDatabase.bas
' (Utilizadas en la funci├│n SaveSolicitudPC)
'******************************************************************************

' ID de la solicitud padre (usado en SaveSolicitudPC)
Public SolicitudID As Long

' N├║mero de expediente (copiado desde solicitud)
Public NumeroExpediente As String

' Tipo de solicitud (copiado desde solicitud)
Public TipoSolicitud As String

' Descripci├│n del cambio propuesto
Public DescripcionCambio As String

' Justificaci├│n del cambio propuesto
Public JustificacionCambio As String

' Impacto en seguridad
Public ImpactoSeguridad As String

' Impacto en calidad
Public ImpactoCalidad As String

' Fecha de ├║ltima modificaci├│n
Public FechaUltimaModificacion As Date

' Estado del registro de datos PC
Public Estado As String

' Indica si el registro est├í activo
Public Activo As Boolean
