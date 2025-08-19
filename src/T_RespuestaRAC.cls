VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_RespuestaRAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

'******************************************************************************
' CLASE: T_RespuestaRAC
' DESCRIPCIÓN: Tipo de datos para respuestas del Sistema RAC
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas basándose en Especificaciones_Integracion_CONDOR.md
'******************************************************************************


'******************************************************************************
' PROPIEDADES PÚBLICAS
' Basadas en la sección 2.4 de Especificaciones_Integracion_CONDOR.md
'******************************************************************************

' Indica si la operación fue exitosa
Public Exitoso As Boolean

' Número de registro RAC asignado
Public NumeroRAC As String

' Fecha de registro en el sistema RAC
Public FechaRegistro As Date

' Estado actual en el sistema RAC
Public EstadoRAC As String

' Mensaje de error en caso de fallo
Public MensajeError As String

' Código de error específico
Public CodigoError As Long

'******************************************************************************
' PROPIEDADES ADICIONALES PARA INTEGRACIÓN COMPLETA
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

' Versión de la API RAC utilizada
Public VersionAPI As String

' Datos adicionales en formato JSON
Public DatosAdicionales As String

' Número de intentos realizados
Public NumeroIntentos As Long

' Indica si requiere acción manual
Public RequiereAccionManual As Boolean

'******************************************************************************
' MÉTODOS PÚBLICOS
'******************************************************************************

' Valida si la respuesta es exitosa y completa
Public Function EsExitoso() As Boolean
    EsExitoso = (Exitoso And Len(NumeroRAC) > 0)
End Function

' Determina si hubo un error en la operación
Public Function TieneError() As Boolean
    TieneError = (Not Exitoso Or _
                  CodigoError <> 0 Or _
                  Len(MensajeError) > 0)
End Function

' Obtiene descripción completa del estado
Public Function ObtenerDescripcionEstado() As String
    If EsExitoso() Then
        ObtenerDescripcionEstado = "Registro exitoso - RAC: " & NumeroRAC & " - Estado: " & EstadoRAC
    Else
        ObtenerDescripcionEstado = "Error " & CodigoError & ": " & MensajeError
    End If
End Function

' Determina si la respuesta es válida para procesamiento
Public Function EsRespuestaValida() As Boolean
    EsRespuestaValida = (Len(SolicitudID) > 0 And _
                        FechaRespuesta > 0 And _
                        Len(URLServicio) > 0)
End Function

' Determina si se puede reintentar la operación
Public Function PuedeReintentar() As Boolean
    ' Se puede reintentar si no es exitoso, no requiere acción manual
    ' y el número de intentos es menor a un límite
    PuedeReintentar = (Not EsExitoso() And _
                      Not RequiereAccionManual And _
                      NumeroIntentos < 5 And _
                      CodigoError <> 401 And _
                      CodigoError <> 403)
End Function











