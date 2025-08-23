Option Compare Database
Option Explicit

#If DEV_MODE Then

' CMockNotificationService.cls
' Mock del servicio de notificaciones para pruebas unitarias
' Simula el comportamiento de CNotificationService sin acceder a bases de datos externas


' Implementa la interfaz INotificationService
Implements INotificationService

' Variables para tracking de llamadas en tests
Public UltimaLlamada_Destinatarios As String
Public UltimaLlamada_Asunto As String
Public UltimaLlamada_CuerpoHTML As String
Public UltimaLlamada_URLAdjunto As String
Public NumeroLlamadas As Long
Public SimularError As Boolean
Public ValorRetorno As Boolean

' Constructor
Private Sub Class_Initialize()
    ' Inicializar valores por defecto
    Me.LimpiarHistorial
    SimularError = False
    ValorRetorno = True
End Sub



' Método para limpiar el historial de llamadas
Public Sub LimpiarHistorial()
    UltimaLlamada_Destinatarios = ""
    UltimaLlamada_Asunto = ""
    UltimaLlamada_CuerpoHTML = ""
    UltimaLlamada_URLAdjunto = ""
    NumeroLlamadas = 0
End Sub

' Implementación mock del método EnviarNotificacion
Private Function INotificationService_EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    
    ' Incrementar contador de llamadas
    NumeroLlamadas = NumeroLlamadas + 1
    
    ' Guardar parámetros de la última llamada
    UltimaLlamada_Destinatarios = destinatarios
    UltimaLlamada_Asunto = asunto
    UltimaLlamada_CuerpoHTML = cuerpoHTML
    UltimaLlamada_URLAdjunto = urlAdjunto
    
    ' Simular error si está configurado
    If SimularError Then
        Err.Raise 9999, "CMockNotificationService", "Error simulado para pruebas"
    End If
    
    ' Retornar valor configurado
    INotificationService_EnviarNotificacion = ValorRetorno
End Function

' Implementación pública del método EnviarNotificacion
Public Function EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    ' Delegar a la implementación de la interfaz
    EnviarNotificacion = INotificationService_EnviarNotificacion(destinatarios, asunto, cuerpoHTML, urlAdjunto)
End Function

' Método helper para verificar si se llamó con parámetros específicos
Public Function FueLlamadoCon( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    
    FueLlamadoCon = (UltimaLlamada_Destinatarios = destinatarios) And _
                    (UltimaLlamada_Asunto = asunto) And _
                    (UltimaLlamada_CuerpoHTML = cuerpoHTML) And _
                    (UltimaLlamada_URLAdjunto = urlAdjunto)
End Function

#End If













