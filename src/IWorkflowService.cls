Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Valida si una transición es permitida en el nuevo flujo de 7 estados
Public Function ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
End Function

' Obtiene los siguientes estados válidos para el flujo refactorizado
Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Object
End Function

' Verifica si un estado es final en el nuevo flujo (Aprobada = ID 7)
Public Function IsEstadoFinal(ByVal estadoActual As String) As Boolean
End Function