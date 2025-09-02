VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Function IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
End Function

Public Function GetNextStates(ByVal idEstadoActual As Long, ByVal usuarioRol As String) As Scripting.Dictionary
End Function




