Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' Clase de tipo de datos para tbLogCambios
' Basada en el Anexo A del CONDOR_MASTER_PLAN.md
' Propósito: Auditoría de cambios realizados en el sistema

' Campos privados según especificación tbLogCambios
Private m_idLogCambio As Long
Private m_FechaHora As Date
Private m_Usuario As String
Private m_tabla As String
Private m_registro As Long
Private m_campo As String
Private m_valorAnterior As String
Private m_valorNuevo As String
Private m_TipoOperacion As String

' Propiedades públicas
Public Property Get IdLogCambio() As Long
    IdLogCambio = m_idLogCambio
End Property

Public Property Let IdLogCambio(ByVal value As Long)
    m_idLogCambio = value
End Property

Public Property Get fechaHora() As Date
    fechaHora = m_FechaHora
End Property

Public Property Let fechaHora(ByVal value As Date)
    m_FechaHora = value
End Property

Public Property Get usuario() As String
    usuario = m_Usuario
End Property

Public Property Let usuario(ByVal value As String)
    m_Usuario = value
End Property

Public Property Get Tabla() As String
    Tabla = m_tabla
End Property

Public Property Let Tabla(ByVal value As String)
    m_tabla = value
End Property

Public Property Get Registro() As Long
    Registro = m_registro
End Property

Public Property Let Registro(ByVal value As Long)
    m_registro = value
End Property

Public Property Get Campo() As String
    Campo = m_campo
End Property

Public Property Let Campo(ByVal value As String)
    m_campo = value
End Property

Public Property Get ValorAnterior() As String
    ValorAnterior = m_valorAnterior
End Property

Public Property Let ValorAnterior(ByVal value As String)
    m_valorAnterior = value
End Property

Public Property Get ValorNuevo() As String
    ValorNuevo = m_valorNuevo
End Property

Public Property Let ValorNuevo(ByVal value As String)
    m_valorNuevo = value
End Property

Public Property Get tipoOperacion() As String
    tipoOperacion = m_TipoOperacion
End Property

Public Property Let tipoOperacion(ByVal value As String)
    m_TipoOperacion = value
End Property