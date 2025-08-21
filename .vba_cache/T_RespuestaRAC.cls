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
' VARIABLES MIEMBRO PRIVADAS
'******************************************************************************
Private m_Exitoso As Boolean
Private m_NumeroRAC As String
Private m_FechaRegistro As Date
Private m_EstadoRAC As String
Private m_MensajeError As String
Private m_CodigoError As Long
Private m_SolicitudID As String
Private m_TiempoRespuesta As Long
Private m_URLServicio As String
Private m_TokenSeguimiento As String
Private m_FechaRespuesta As Date
Private m_VersionAPI As String
Private m_DatosAdicionales As String
Private m_NumeroIntentos As Long
Private m_RequiereAccionManual As Boolean

'******************************************************************************
' PROPIEDADES PÚBLICAS (GET/LET)
'******************************************************************************

Public Property Get Exitoso() As Boolean
    Exitoso = m_Exitoso
End Property
Public Property Let Exitoso(ByVal value As Boolean)
    m_Exitoso = value
End Property

Public Property Get NumeroRAC() As String
    NumeroRAC = m_NumeroRAC
End Property
Public Property Let NumeroRAC(ByVal value As String)
    m_NumeroRAC = value
End Property

Public Property Get FechaRegistro() As Date
    FechaRegistro = m_FechaRegistro
End Property
Public Property Let FechaRegistro(ByVal value As Date)
    m_FechaRegistro = value
End Property

Public Property Get EstadoRAC() As String
    EstadoRAC = m_EstadoRAC
End Property
Public Property Let EstadoRAC(ByVal value As String)
    m_EstadoRAC = value
End Property

Public Property Get MensajeError() As String
    MensajeError = m_MensajeError
End Property
Public Property Let MensajeError(ByVal value As String)
    m_MensajeError = value
End Property

Public Property Get CodigoError() As Long
    CodigoError = m_CodigoError
End Property
Public Property Let CodigoError(ByVal value As Long)
    m_CodigoError = value
End Property

Public Property Get SolicitudID() As String
    SolicitudID = m_SolicitudID
End Property
Public Property Let SolicitudID(ByVal value As String)
    m_SolicitudID = value
End Property

Public Property Get TiempoRespuesta() As Long
    TiempoRespuesta = m_TiempoRespuesta
End Property
Public Property Let TiempoRespuesta(ByVal value As Long)
    m_TiempoRespuesta = value
End Property

Public Property Get URLServicio() As String
    URLServicio = m_URLServicio
End Property
Public Property Let URLServicio(ByVal value As String)
    m_URLServicio = value
End Property

Public Property Get TokenSeguimiento() As String
    TokenSeguimiento = m_TokenSeguimiento
End Property
Public Property Let TokenSeguimiento(ByVal value As String)
    m_TokenSeguimiento = value
End Property

Public Property Get FechaRespuesta() As Date
    FechaRespuesta = m_FechaRespuesta
End Property
Public Property Let FechaRespuesta(ByVal value As Date)
    m_FechaRespuesta = value
End Property

Public Property Get VersionAPI() As String
    VersionAPI = m_VersionAPI
End Property
Public Property Let VersionAPI(ByVal value As String)
    m_VersionAPI = value
End Property

Public Property Get DatosAdicionales() As String
    DatosAdicionales = m_DatosAdicionales
End Property
Public Property Let DatosAdicionales(ByVal value As String)
    m_DatosAdicionales = value
End Property

Public Property Get NumeroIntentos() As Long
    NumeroIntentos = m_NumeroIntentos
End Property
Public Property Let NumeroIntentos(ByVal value As Long)
    m_NumeroIntentos = value
End Property

Public Property Get RequiereAccionManual() As Boolean
    RequiereAccionManual = m_RequiereAccionManual
End Property
Public Property Let RequiereAccionManual(ByVal value As Boolean)
    m_RequiereAccionManual = value
End Property

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













