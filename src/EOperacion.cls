VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' Clase de tipo de datos para tb_Operaciones_Log
' Basada en el Anexo A del CONDOR_MASTER_PLAN.md
' Propósito: Registro de operaciones clave del sistema para auditoría y trazabilidad

' Campos privados según especificación tb_Operaciones_Log
Private m_idOperacion As Long
Private m_FechaHora As Date
Private m_Usuario As String
Private m_TipoOperacion As String
Private m_entidad As String
Private m_idEntidad As Long
Private m_descripcion As String
Private m_resultado As String
Private m_Detalles As String

' Propiedades públicas
Public Property Get IdOperacion() As Long
    IdOperacion = m_idOperacion
End Property

Public Property Let IdOperacion(ByVal value As Long)
    m_idOperacion = value
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

Public Property Get tipoOperacion() As String
    tipoOperacion = m_TipoOperacion
End Property

Public Property Let tipoOperacion(ByVal value As String)
    m_TipoOperacion = value
End Property

Public Property Get Entidad() As String
    Entidad = m_entidad
End Property

Public Property Let Entidad(ByVal value As String)
    m_entidad = value
End Property

Public Property Get IdEntidad() As Long
    IdEntidad = m_idEntidad
End Property

Public Property Let IdEntidad(ByVal value As Long)
    m_idEntidad = value
End Property

Public Property Get Descripcion() As String
    Descripcion = m_descripcion
End Property

Public Property Let Descripcion(ByVal value As String)
    m_descripcion = value
End Property

Public Property Get Resultado() As String
    Resultado = m_resultado
End Property

Public Property Let Resultado(ByVal value As String)
    m_resultado = value
End Property

Public Property Get detalles() As String
    detalles = m_Detalles
End Property

Public Property Let detalles(ByVal value As String)
    m_Detalles = value
End Property



