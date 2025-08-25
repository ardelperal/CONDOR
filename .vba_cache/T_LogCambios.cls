Option Compare Database
Option Explicit

' Clase de tipo de datos para tbLogCambios
' Basada en el Anexo A del CONDOR_MASTER_PLAN.md
' Propósito: Auditoría de cambios realizados en el sistema

' Campos privados según especificación tbLogCambios
Private m_idLogCambio As Long
Private m_fechaHora As Date
Private m_usuario As String
Private m_tabla As String
Private m_registro As Long
Private m_campo As String
Private m_valorAnterior As String
Private m_valorNuevo As String
Private m_tipoOperacion As String

' Propiedades públicas
Public Property Get IdLogCambio() As Long
    IdLogCambio = m_idLogCambio
End Property

Public Property Let IdLogCambio(ByVal value As Long)
    m_idLogCambio = value
End Property

Public Property Get FechaHora() As Date
    FechaHora = m_fechaHora
End Property

Public Property Let FechaHora(ByVal value As Date)
    m_fechaHora = value
End Property

Public Property Get Usuario() As String
    Usuario = m_usuario
End Property

Public Property Let Usuario(ByVal value As String)
    m_usuario = value
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

Public Property Get TipoOperacion() As String
    TipoOperacion = m_tipoOperacion
End Property

Public Property Let TipoOperacion(ByVal value As String)
    m_tipoOperacion = value
End Property