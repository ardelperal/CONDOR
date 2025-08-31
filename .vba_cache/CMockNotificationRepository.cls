VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockNotificationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements INotificationRepository

' Propiedades públicas para verificación en pruebas
Public WasEnqueueEmailCalled As Boolean
Public LastDestinatarios As String
Public LastAsunto As String
Public LastCuerpoHTML As String
Public LastUsuario As String
Public LastAdminEmail As String
Public LastAdjuntoPath As String
Public EnqueueEmailReturnValue As Boolean

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' Reinicia el estado del mock
Public Sub Reset()
    WasEnqueueEmailCalled = False
    LastDestinatarios = ""
    LastAsunto = ""
    LastCuerpoHTML = ""
    LastUsuario = ""
    LastAdminEmail = ""
    LastAdjuntoPath = ""
    EnqueueEmailReturnValue = True
End Sub

' Implementación de la interfaz
Private Function INotificationRepository_EnqueueEmail(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpoHTML As String, ByVal usuario As String, ByVal adminEmail As String, Optional ByVal adjuntoPath As String = "") As Boolean
    ' Registrar la llamada para verificación
    WasEnqueueEmailCalled = True
    LastDestinatarios = destinatarios
    LastAsunto = asunto
    LastCuerpoHTML = cuerpoHTML
    LastUsuario = usuario
    LastAdminEmail = adminEmail
    LastAdjuntoPath = adjuntoPath
    
    ' Devolver el valor configurado
    INotificationRepository_EnqueueEmail = EnqueueEmailReturnValue
End Function

' Método público de conveniencia
Public Function EnqueueEmail(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpoHTML As String, ByVal usuario As String, ByVal adminEmail As String, Optional ByVal adjuntoPath As String = "") As Boolean
    EnqueueEmail = INotificationRepository_EnqueueEmail(destinatarios, asunto, cuerpoHTML, usuario, adminEmail, adjuntoPath)
End Function