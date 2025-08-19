VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

Public Property Get idSolicitud() As Long
End Property

Public Property Let idSolicitud(ByVal RHS As Long)
End Property

Public Property Get idExpediente() As String
End Property

Public Property Let idExpediente(ByVal RHS As String)
End Property

Public Property Get tipoSolicitud() As String
End Property

Public Property Let tipoSolicitud(ByVal RHS As String)
End Property

Public Property Get codigoSolicitud() As String
End Property

Public Property Let codigoSolicitud(ByVal RHS As String)
End Property

Public Property Get estadoInterno() As String
End Property

Public Property Let estadoInterno(ByVal RHS As String)
End Property

Public Function Load(ByVal ID As Long) As Boolean
End Function

Public Function Save() As Boolean
End Function

Public Function ChangeState(ByVal newState As String) As Boolean
End Function










