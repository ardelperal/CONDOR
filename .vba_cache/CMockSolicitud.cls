Option Compare Database
Option Explicit
' ============================================================================
' Clase: CMockSolicitud
' Descripción: Mock de una solicitud para pruebas unitarias.
' ============================================================================

Implements ISolicitud

' --- PROPIEDADES PÚBLICAS PARA CONFIGURAR EL MOCK ---
Public idSolicitud As Long
Public idExpediente As String
Public tipoSolicitud As String
Public codigoSolicitud As String
Public estadoInterno As String
Public Datos As Object

' --- IMPLEMENTACIÓN DE LA INTERFAZ ISolicitud ---

Private Property Get ISolicitud_idSolicitud() As Long
    ISolicitud_idSolicitud = idSolicitud
End Property

Private Property Let ISolicitud_idSolicitud(ByVal RHS As Long)
    idSolicitud = RHS
End Property

Private Property Get ISolicitud_idExpediente() As String
    ISolicitud_idExpediente = idExpediente
End Property

Private Property Let ISolicitud_idExpediente(ByVal RHS As String)
    idExpediente = RHS
End Property

Private Property Get ISolicitud_tipoSolicitud() As String
    ISolicitud_tipoSolicitud = tipoSolicitud
End Property

Private Property Let ISolicitud_tipoSolicitud(ByVal RHS As String)
    tipoSolicitud = RHS
End Property

Private Property Get ISolicitud_codigoSolicitud() As String
    ISolicitud_codigoSolicitud = codigoSolicitud
End Property

Private Property Let ISolicitud_codigoSolicitud(ByVal RHS As String)
    codigoSolicitud = RHS
End Property

Private Property Get ISolicitud_estadoInterno() As String
    ISolicitud_estadoInterno = estadoInterno
End Property

Private Property Let ISolicitud_estadoInterno(ByVal RHS As String)
    estadoInterno = RHS
End Property

Private Property Get ISolicitud_Datos() As Object
    Set ISolicitud_Datos = Datos
End Property

' --- MÉTODOS DE LA INTERFAZ (NO NECESARIOS PARA ESTE MOCK) ---

Private Function ISolicitud_Load(ByVal ID As Long) As Boolean
    ' No implementado para este mock
End Function

Private Function ISolicitud_Save() As Boolean
    ' No implementado para este mock
End Function

Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean
    ' No implementado para este mock
End Function
