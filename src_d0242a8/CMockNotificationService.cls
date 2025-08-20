VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockNotificationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CMockNotificationService.cls
' Mock del servicio de notificaciones para pruebas unitarias
' Simula el comportamiento de CNotificationService sin acceder a bases de datos externas

Option Compare Database
Option Explicit

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

' M├®todo para limpiar el historial de llamadas
Public Sub LimpiarHistorial()
    UltimaLlamada_Destinatarios = ""
    UltimaLlamada_Asunto = ""
    UltimaLlamada_CuerpoHTML = ""
    UltimaLlamada_URLAdjunto = ""
    NumeroLlamadas = 0
End Sub

' Implementaci├│n mock del m├®todo EnviarNotificacion
Private Function INotificationService_EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    
    ' Incrementar contador de llamadas
    NumeroLlamadas = NumeroLlamadas + 1
    
    ' Guardar par├ímetros de la ├║ltima llamada
    UltimaLlamada_Destinatarios = destinatarios
    UltimaLlamada_Asunto = asunto
    UltimaLlamada_CuerpoHTML = cuerpoHTML
    UltimaLlamada_URLAdjunto = urlAdjunto
    
    ' Simular error si est├í configurado
    If SimularError Then
        Err.Raise 9999, "CMockNotificationService", "Error simulado para pruebas"
    End If
    
    ' Retornar valor configurado
    INotificationService_EnviarNotificacion = ValorRetorno
End Function

' M├®todo helper para verificar si se llam├│ con par├ímetros espec├¡ficos
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
