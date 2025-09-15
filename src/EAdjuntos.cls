Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' Clase de tipo de datos para tbAdjuntos
' Basada en el Anexo A del CONDOR_MASTER_PLAN.md
' Propósito: Gestión de ficheros adjuntos a las solicitudes

' Campos privados según especificación tbAdjuntos
Private m_idAdjunto As Long
Private m_idSolicitud As Long
Private m_nombreArchivo As String
Private m_fechaSubida As Date
Private m_usuarioSubida As String
Private m_Descripcion As String

' Propiedades públicas
Public Property Get IdAdjunto() As Long
    IdAdjunto = m_idAdjunto
End Property

Public Property Let IdAdjunto(ByVal value As Long)
    m_idAdjunto = value
End Property

Public Property Get idSolicitud() As Long
    idSolicitud = m_idSolicitud
End Property

Public Property Let idSolicitud(ByVal value As Long)
    m_idSolicitud = value
End Property

Public Property Get NombreArchivo() As String
    NombreArchivo = m_nombreArchivo
End Property

Public Property Let NombreArchivo(ByVal value As String)
    m_nombreArchivo = value
End Property

Public Property Get FechaSubida() As Date
    FechaSubida = m_fechaSubida
End Property

Public Property Let FechaSubida(ByVal value As Date)
    m_fechaSubida = value
End Property

Public Property Get UsuarioSubida() As String
    UsuarioSubida = m_usuarioSubida
End Property

Public Property Let UsuarioSubida(ByVal value As String)
    m_usuarioSubida = value
End Property

Public Property Get descripcion() As String
    descripcion = m_Descripcion
End Property

Public Property Let descripcion(ByVal value As String)
    m_Descripcion = value
End Property