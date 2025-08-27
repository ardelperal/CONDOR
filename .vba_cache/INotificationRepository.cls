VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotificationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'******************************************************************************
' INTERFAZ: INotificationRepository
' Propósito: Define el contrato para el repositorio de notificaciones
' Patrón: Repository Pattern para separación de responsabilidades
' Versión: 1.0
'******************************************************************************

' Método para encolar un correo en la base de datos
' El repositorio es responsable de generar el ID
Public Function EncolarCorreo(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpo As String, ByVal cc As String, ByVal bcc As String, ByVal adjunto As String) As Boolean
End Function   ' Implementación debe ser proporcionada por las clases concretas
