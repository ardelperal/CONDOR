VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_RespuestaRAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: T_RespuestaRAC
' DESCRIPCI├ôN: Tipo de datos para respuestas del Sistema RAC
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas bas├índose en Especificaciones_Integracion_CONDOR.md
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' PROPIEDADES P├ÜBLICAS
' Basadas en la secci├│n 2.4 de Especificaciones_Integracion_CONDOR.md
'******************************************************************************

' Indica si la operaci├│n fue exitosa
Public Exitoso As Boolean

' N├║mero de registro RAC asignado
Public NumeroRAC As String

' Fecha de registro en el sistema RAC
Public FechaRegistro As Date

' Estado actual en el sistema RAC
Public EstadoRAC As String

' Mensaje de error en caso de fallo
Public MensajeError As String

' C├│digo de error espec├¡fico
Public CodigoError As Long

'******************************************************************************
' PROPIEDADES ADICIONALES PARA INTEGRACI├ôN COMPLETA
'******************************************************************************

' ID de la solicitud asociada
Public SolicitudID As String

' Tiempo de respuesta del servicio RAC (en milisegundos)
Public TiempoRespuesta As Long

' URL del servicio RAC utilizado
Public URLServicio As String

' Token de seguimiento para consultas posteriores
Public TokenSeguimiento As String

' Fecha y hora de la respuesta
Public FechaRespuesta As Date

' Versi├│n de la API RAC utilizada
Public VersionAPI As String

' Datos adicionales en formato JSON
Public DatosAdicionales As String

' N├║mero de intentos realizados
Public NumeroIntentos As Long

' Indica si requiere acci├│n manual
Public RequiereAccionManual As Boolean

'******************************************************************************
' M├ëTODOS P├ÜBLICOS
'******************************************************************************

' Valida si la respuesta es exitosa y completa
Public Function EsExitoso() As Boolean
    EsExitoso = (Exitoso And Len(NumeroRAC) > 0)
End Function

' Determina si hubo un error en la operaci├│n
Public Function TieneError() As Boolean
    TieneError = (Not Exitoso Or _
                  CodigoError <> 0 Or _
                  Len(MensajeError) > 0)
End Function

' Obtiene descripci├│n completa del estado
Public Function ObtenerDescripcionEstado() As String
    If EsExitoso() Then
        ObtenerDescripcionEstado = "Registro exitoso - RAC: " & NumeroRAC & " - Estado: " & EstadoRAC
    Else
        ObtenerDescripcionEstado = "Error " & CodigoError & ": " & MensajeError
    End If
End Function

' Determina si la respuesta es v├ílida para procesamiento
Public Function EsRespuestaValida() As Boolean
    EsRespuestaValida = (Len(SolicitudID) > 0 And _
                        FechaRespuesta > 0 And _
                        Len(URLServicio) > 0)
End Function

' Determina si se puede reintentar la operaci├│n
Public Function PuedeReintentar() As Boolean
    ' Se puede reintentar si no es exitoso, no requiere acci├│n manual
    ' y el n├║mero de intentos es menor a un l├¡mite
    PuedeReintentar = (Not EsExitoso() And _
                      Not RequiereAccionManual And _
                      NumeroIntentos < 5 And _
                      CodigoError <> 401 And _
                      CodigoError <> 403)
End Function
