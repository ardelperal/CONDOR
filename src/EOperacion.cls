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
Private m_fechaHora As Date
Private m_usuario As String
Private m_tipoOperacion As String
Private m_entidad As String
Private m_idEntidad As Long
Private m_descripcion As String
Private m_resultado As String
Private m_detalles As String

' Propiedades públicas
Public Property Get IdOperacion() As Long
    IdOperacion = m_idOperacion
End Property

Public Property Let IdOperacion(ByVal value As Long)
    m_idOperacion = value
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

Public Property Get TipoOperacion() As String
    TipoOperacion = m_tipoOperacion
End Property

Public Property Let TipoOperacion(ByVal value As String)
    m_tipoOperacion = value
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

Public Property Get Detalles() As String
    Detalles = m_detalles
End Property

Public Property Let Detalles(ByVal value As String)
    m_detalles = value
End Property

