VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNotificationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ===================================================================
' Clase: CNotificationService
' Descripción: Implementación concreta del servicio de notificaciones por correo electrónico
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================

Option Explicit

' Implementa la interfaz INotificationService
Implements INotificationService

' ===================================================================
' IMPLEMENTACIÓN DE MÉTODOS DE NOTIFICACIÓN PRINCIPAL
' ===================================================================

Private Function INotificationService_EnviarNotificacion(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String, ByRef destinatario As String, ByRef asunto As String, ByRef cuerpo As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim mensajeError As String
    
    ' Validar parámetros de entrada
    If Not INotificationService_ValidarParametrosNotificacion(solicitud, tipoEvento, mensajeError) Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Construir componentes del correo
    destinatario = INotificationService_ObtenerDestinatario(solicitud, tipoEvento)
    asunto = INotificationService_ConstruirAsunto(solicitud, tipoEvento)
    cuerpo = INotificationService_ConstruirCuerpo(solicitud, tipoEvento)
    
    ' TODO: Aquí se implementaría el envío real del correo
    ' Por ahora, simulamos el éxito si todos los componentes están construidos
    If Len(destinatario) > 0 And Len(asunto) > 0 And Len(cuerpo) > 0 Then
        INotificationService_EnviarNotificacion = True
    Else
        INotificationService_EnviarNotificacion = False
    End If
    
    Exit Function
    
ErrorHandler:
    INotificationService_EnviarNotificacion = False
End Function

' ===================================================================
' IMPLEMENTACIÓN DE MÉTODOS DE CONSTRUCCIÓN DE CONTENIDO
' ===================================================================

Private Function INotificationService_ConstruirAsunto(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String) As String
    On Error GoTo ErrorHandler
    
    Dim asunto As String
    
    Select Case tipoEvento
        Case "CambioEstado"
            asunto = "CONDOR - Cambio de Estado: " & solicitud.NumeroExpediente & " - " & solicitud.Estado
        Case "Creacion"
            asunto = "CONDOR - Nueva Solicitud: " & solicitud.NumeroExpediente & " - " & solicitud.TipoSolicitud
        Case "Aprobacion"
            asunto = "CONDOR - Solicitud Aprobada: " & solicitud.NumeroExpediente
        Case "Rechazo"
            asunto = "CONDOR - Solicitud Rechazada: " & solicitud.NumeroExpediente
        Case Else
            asunto = "CONDOR - Notificación: " & solicitud.NumeroExpediente
    End Select
    
    INotificationService_ConstruirAsunto = asunto
    Exit Function
    
ErrorHandler:
    INotificationService_ConstruirAsunto = ""
End Function

Private Function INotificationService_ConstruirCuerpo(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String) As String
    On Error GoTo ErrorHandler
    
    Dim cuerpo As String
    
    cuerpo = "Estimado usuario," & vbCrLf & vbCrLf
    
    Select Case tipoEvento
        Case "CambioEstado"
            cuerpo = cuerpo & "Le informamos que la solicitud " & solicitud.NumeroExpediente & " ha cambiado de estado." & vbCrLf & vbCrLf
            cuerpo = cuerpo & "Nuevo estado: " & solicitud.Estado & vbCrLf
        Case "Creacion"
            cuerpo = cuerpo & "Se ha creado una nueva solicitud que requiere su atención." & vbCrLf & vbCrLf
        Case "Aprobacion"
            cuerpo = cuerpo & "Su solicitud " & solicitud.NumeroExpediente & " ha sido aprobada." & vbCrLf & vbCrLf
        Case "Rechazo"
            cuerpo = cuerpo & "Su solicitud " & solicitud.NumeroExpediente & " ha sido rechazada." & vbCrLf & vbCrLf
        Case Else
            cuerpo = cuerpo & "Hay una actualización en su solicitud " & solicitud.NumeroExpediente & "." & vbCrLf & vbCrLf
    End Select
    
    ' Agregar detalles de la solicitud
    cuerpo = cuerpo & "Detalles de la solicitud:" & vbCrLf
    cuerpo = cuerpo & "- Número de Expediente: " & solicitud.NumeroExpediente & vbCrLf
    cuerpo = cuerpo & "- Tipo: " & solicitud.TipoSolicitud & vbCrLf
    cuerpo = cuerpo & "- Descripción: " & solicitud.Descripcion & vbCrLf
    cuerpo = cuerpo & "- Estado Actual: " & solicitud.Estado & vbCrLf
    cuerpo = cuerpo & "- Fecha de Creación: " & Format(solicitud.FechaCreacion, "dd/mm/yyyy hh:mm") & vbCrLf & vbCrLf
    
    cuerpo = cuerpo & "Saludos cordiales," & vbCrLf
    cuerpo = cuerpo & "Sistema CONDOR"
    
    INotificationService_ConstruirCuerpo = cuerpo
    Exit Function
    
ErrorHandler:
    INotificationService_ConstruirCuerpo = ""
End Function

Private Function INotificationService_ObtenerDestinatario(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String) As String
    On Error GoTo ErrorHandler
    
    Dim destinatario As String
    
    Select Case tipoEvento
        Case "CambioEstado", "Aprobacion", "Rechazo"
            ' Notificar al creador de la solicitud
            destinatario = solicitud.UsuarioCreador
        Case "Creacion"
            ' Notificar a los supervisores (por ahora, usar el creador como fallback)
            destinatario = solicitud.UsuarioCreador
        Case Else
            destinatario = solicitud.UsuarioCreador
    End Select
    
    INotificationService_ObtenerDestinatario = destinatario
    Exit Function
    
ErrorHandler:
    INotificationService_ObtenerDestinatario = ""
End Function

' ===================================================================
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN
' ===================================================================

Private Function INotificationService_ValidarParametrosNotificacion(ByVal solicitud As T_Solicitud, ByVal tipoEvento As String, ByRef mensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    mensajeError = ""
    
    ' Validar que la solicitud no sea nula
    If solicitud Is Nothing Then
        mensajeError = "La solicitud no puede ser nula"
        INotificationService_ValidarParametrosNotificacion = False
        Exit Function
    End If
    
    ' Validar que el tipo de evento no esté vacío
    If Len(Trim(tipoEvento)) = 0 Then
        mensajeError = "El tipo de evento no puede estar vacío"
        INotificationService_ValidarParametrosNotificacion = False
        Exit Function
    End If
    
    ' Validar que la solicitud tenga los campos mínimos requeridos
    If Len(Trim(solicitud.NumeroExpediente)) = 0 Then
        mensajeError = "La solicitud debe tener un número de expediente"
        INotificationService_ValidarParametrosNotificacion = False
        Exit Function
    End If
    
    If Len(Trim(solicitud.UsuarioCreador)) = 0 Then
        mensajeError = "La solicitud debe tener un usuario creador"
        INotificationService_ValidarParametrosNotificacion = False
        Exit Function
    End If
    
    ' Validar formato de email del usuario creador
    If InStr(solicitud.UsuarioCreador, "@") = 0 Then
        mensajeError = "El usuario creador debe tener un formato de email válido"
        INotificationService_ValidarParametrosNotificacion = False
        Exit Function
    End If
    
    INotificationService_ValidarParametrosNotificacion = True
    Exit Function
    
ErrorHandler:
    mensajeError = "Error en validación: " & Err.Description
    INotificationService_ValidarParametrosNotificacion = False
End Function