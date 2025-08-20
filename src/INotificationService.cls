Option Compare Database
Option Explicit
' INotificationService.cls
' Interfaz para el servicio de notificaciones asíncrono
' Basado en las Especificaciones de Integración - Sección 2.5
' El servicio encola correos en la base de datos Correos_datos.accdb


' Método principal para enviar notificaciones
' Encola un correo en la base de datos externa para procesamiento asíncrono
' Parámetros:
'   destinatarios: Lista de correos separados por punto y coma
'   asunto: Asunto del correo
'   cuerpoHTML: Cuerpo del correo en formato HTML
'   urlAdjunto: Ruta completa al archivo adjunto (opcional)
' Retorna: True si se encoló correctamente, False en caso contrario
Public Function EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
End Function









