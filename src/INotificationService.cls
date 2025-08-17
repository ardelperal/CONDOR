VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotificationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ===================================================================
' Interfaz: INotificationService
' Descripción: Contrato para servicios de notificaciones por correo electrónico
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================

Option Explicit

' ===================================================================
' MÉTODOS DE NOTIFICACIÓN PRINCIPAL
' ===================================================================

' Envía una notificación por correo electrónico para eventos de solicitudes
' Parámetros:
'   solicitud: Objeto T_Solicitud que genera el evento
'   tipoEvento: Tipo de evento ("CambioEstado", "Creacion", "Aprobacion", etc.)
'   destinatario: Variable de salida con el email del destinatario
'   asunto: Variable de salida con el asunto del correo
'   cuerpo: Variable de salida con el cuerpo del mensaje
' Retorna: True si la notificación se procesó correctamente, False en caso contrario
Public Function EnviarNotificacion(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String, ByRef destinatario As String, ByRef asunto As String, ByRef cuerpo As String) As Boolean
End Function

' ===================================================================
' MÉTODOS DE CONSTRUCCIÓN DE CONTENIDO
' ===================================================================

' Construye el asunto del correo basado en la solicitud y tipo de evento
' Parámetros:
'   solicitud: Objeto T_Solicitud
'   tipoEvento: Tipo de evento
' Retorna: String con el asunto formateado
Public Function ConstruirAsunto(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String) As String
End Function

' Construye el cuerpo del mensaje basado en la solicitud y tipo de evento
' Parámetros:
'   solicitud: Objeto T_Solicitud
'   tipoEvento: Tipo de evento
' Retorna: String con el cuerpo del mensaje formateado
Public Function ConstruirCuerpo(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String) As String
End Function

' Obtiene el destinatario apropiado basado en la solicitud y tipo de evento
' Parámetros:
'   solicitud: Objeto T_Solicitud
'   tipoEvento: Tipo de evento
' Retorna: String con el email del destinatario
Public Function ObtenerDestinatario(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String) As String
End Function

' ===================================================================
' MÉTODOS DE VALIDACIÓN
' ===================================================================

' Valida que los parámetros de notificación sean válidos
' Parámetros:
'   solicitud: Objeto T_Solicitud a validar
'   tipoEvento: Tipo de evento a validar
'   mensajeError: Variable de salida con el mensaje de error (si aplica)
' Retorna: True si los parámetros son válidos, False en caso contrario
Public Function ValidarParametrosNotificacion(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String, ByRef mensajeError As String) As Boolean
End Function