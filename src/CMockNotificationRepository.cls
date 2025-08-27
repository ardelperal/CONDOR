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

'******************************************************************************
' CLASE MOCK: CMockNotificationRepository
' Propósito: Mock del repositorio de notificaciones para pruebas unitarias
' Patrón: Test Double (Mock Object)
' Versión: 1.0
'******************************************************************************

' Propiedades públicas para verificación en pruebas
Public WasCalled As Boolean
Public LastAsunto As String
Public LastCuerpo As String
Public LastDestinatarios As String
Public LastCc As String
Public LastBcc As String
Public LastAdjunto As String
Public CallCount As Integer

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' Método para resetear el estado del mock
Public Sub Reset()
    WasCalled = False
    LastAsunto = ""
    LastCuerpo = ""
    LastDestinatarios = ""
    LastCc = ""
    LastBcc = ""
    LastAdjunto = ""
    CallCount = 0
End Sub

' Implementación del método de la interfaz
Private Function INotificationRepository_EncolarCorreo(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpo As String, ByVal cc As String, ByVal bcc As String, ByVal adjunto As String) As Boolean
    ' Registrar la llamada para verificación en pruebas
    LastAsunto = asunto
    LastCuerpo = cuerpo
    LastDestinatarios = destinatarios
    LastCC = cc
    LastBCC = bcc
    LastAdjunto = adjunto
    CallCount = CallCount + 1
    
    INotificationRepository_EncolarCorreo = True
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para EncolarCorreo
Public Function EncolarCorreo(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpo As String, ByVal cc As String, ByVal bcc As String, ByVal adjunto As String) As Boolean
    EncolarCorreo = INotificationRepository_EncolarCorreo(destinatarios, asunto, cuerpo, cc, bcc, adjunto)
End Function