VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Function ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
End Function

Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Scripting.Dictionary
End Function



