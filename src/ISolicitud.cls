Option Compare Database
Option Explicit

Public Property Get IdSolicitud() As Long
End Property

Public Property Let IdSolicitud(ByVal RHS As Long)
End Property

Public Property Get IdExpediente() As String
End Property

Public Property Let IdExpediente(ByVal RHS As String)
End Property

Public Property Get TipoSolicitud() As String
End Property

Public Property Let TipoSolicitud(ByVal RHS As String)
End Property

Public Property Get CodigoSolicitud() As String
End Property

Public Property Let CodigoSolicitud(ByVal RHS As String)
End Property

Public Property Get EstadoInterno() As String
End Property

Public Property Let EstadoInterno(ByVal RHS As String)
End Property

Public Function Load(ByVal id As Long) As Boolean
End Function

Public Function Save() As Boolean
End Function

Public Function ChangeState(ByVal newState As String) As Boolean
End Function


