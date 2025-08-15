Option Compare Database
Option Explicit

Implements ISolicitud

Private m_DatosPC As T_Datos_PC
Private m_IdSolicitud As Long
Private m_IdExpediente As String
Private m_TipoSolicitud As String
Private m_CodigoSolicitud As String
Private m_EstadoInterno As String

' --- 1. Implementación PRIVADA de la Interfaz ---
'     Satisface el contrato "Implements".

Private Property Get ISolicitud_IdSolicitud() As Long
    ISolicitud_IdSolicitud = m_IdSolicitud
End Property

Private Property Let ISolicitud_IdSolicitud(ByVal RHS As Long)
    m_IdSolicitud = RHS
End Property

Private Property Get ISolicitud_IdExpediente() As String
    ISolicitud_IdExpediente = m_IdExpediente
End Property

Private Property Let ISolicitud_IdExpediente(ByVal RHS As String)
    m_IdExpediente = RHS
End Property

Private Property Get ISolicitud_TipoSolicitud() As String
    ISolicitud_TipoSolicitud = m_TipoSolicitud
End Property

Private Property Let ISolicitud_TipoSolicitud(ByVal RHS As String)
    m_TipoSolicitud = RHS
End Property

Private Property Get ISolicitud_CodigoSolicitud() As String
    ISolicitud_CodigoSolicitud = m_CodigoSolicitud
End Property

Private Property Let ISolicitud_CodigoSolicitud(ByVal RHS As String)
    m_CodigoSolicitud = RHS
End Property

Private Property Get ISolicitud_EstadoInterno() As String
    ISolicitud_EstadoInterno = m_EstadoInterno
End Property

Private Property Let ISolicitud_EstadoInterno(ByVal RHS As String)
    m_EstadoInterno = RHS
End Property

Private Function ISolicitud_Load(ByVal id As Long) As Boolean
    ' TODO: Implementar lógica de carga
    ISolicitud_Load = True
End Function

Private Function ISolicitud_Save() As Boolean
    ' TODO: Implementar lógica de guardado
    ISolicitud_Save = True
End Function

Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean
    ' TODO: Implementar lógica de cambio de estado
    ISolicitud_ChangeState = True
End Function

' --- 2. Miembros PÚBLICOS de la Clase ---
'     Exponen la funcionalidad al mundo. Delegan en la implementación de la interfaz.

Public Property Get IdSolicitud() As Long
    IdSolicitud = ISolicitud_IdSolicitud
End Property

Public Property Let IdSolicitud(ByVal val As Long)
    ISolicitud_IdSolicitud = val
End Property

Public Property Get IdExpediente() As String
    IdExpediente = ISolicitud_IdExpediente
End Property

Public Property Let IdExpediente(ByVal val As String)
    ISolicitud_IdExpediente = val
End Property

Public Property Get TipoSolicitud() As String
    TipoSolicitud = ISolicitud_TipoSolicitud
End Property

Public Property Let TipoSolicitud(ByVal val As String)
    ISolicitud_TipoSolicitud = val
End Property

Public Property Get CodigoSolicitud() As String
    CodigoSolicitud = ISolicitud_CodigoSolicitud
End Property

Public Property Let CodigoSolicitud(ByVal val As String)
    ISolicitud_CodigoSolicitud = val
End Property

Public Property Get EstadoInterno() As String
    EstadoInterno = ISolicitud_EstadoInterno
End Property

Public Property Let EstadoInterno(ByVal val As String)
    ISolicitud_EstadoInterno = val
End Property

Public Function Load(ByVal id As Long) As Boolean
    Load = ISolicitud_Load(id)
End Function

Public Function Save() As Boolean
    Save = ISolicitud_Save
End Function

Public Function ChangeState(ByVal newState As String) As Boolean
    ChangeState = ISolicitud_ChangeState(newState)
End Function

' --- Propiedades específicas de PC ---
Public Property Get DatosPC() As T_Datos_PC
    DatosPC = m_DatosPC
End Property

Public Property Let DatosPC(ByRef val As T_Datos_PC)
    m_DatosPC = val
End Property



