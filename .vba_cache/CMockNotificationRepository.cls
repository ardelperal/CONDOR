Attribute VB_Name = "CMockNotificationRepository"
Option Compare Database
Option Explicit
Implements INotificationRepository

Private Function INotificationRepository_EncolarCorreo(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpo As String, ByVal cc As String, ByVal bcc As String, ByVal adjunto As String) As Boolean
    ' Mock: Siempre devuelve True
    INotificationRepository_EncolarCorreo = True
End Function