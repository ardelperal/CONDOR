VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EOperationLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: EOperationLog
' Descripción: Entidad que representa un registro de operación en el sistema
' Arquitectura: Capa de Entidades - Modelo de datos para logs de operaciones
' ============================================================================

' Propiedades de la entidad OperationLog
Private m_FechaHora As Date
Private m_Usuario As String
Private m_TipoOperacion As String
Private m_IDEntidadAfectada As String
Private m_Detalles As String

' Propiedades públicas
Public Property Get FechaHora() As Date
    FechaHora = m_FechaHora
End Property

Public Property Let FechaHora(ByVal value As Date)
    m_FechaHora = value
End Property

Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property

Public Property Let Usuario(ByVal value As String)
    m_Usuario = value
End Property

Public Property Get TipoOperacion() As String
    TipoOperacion = m_TipoOperacion
End Property

Public Property Let TipoOperacion(ByVal value As String)
    m_TipoOperacion = value
End Property

Public Property Get IDEntidadAfectada() As String
    IDEntidadAfectada = m_IDEntidadAfectada
End Property

Public Property Let IDEntidadAfectada(ByVal value As String)
    m_IDEntidadAfectada = value
End Property

Public Property Get Detalles() As String
    Detalles = m_Detalles
End Property

Public Property Let Detalles(ByVal value As String)
    m_Detalles = value
End Property

' Constructor
Public Sub Initialize(ByVal fechaHora As Date, ByVal usuario As String, ByVal tipoOperacion As String, ByVal idEntidadAfectada As String, ByVal detalles As String)
    m_FechaHora = fechaHora
    m_Usuario = usuario
    m_TipoOperacion = tipoOperacion
    m_IDEntidadAfectada = idEntidadAfectada
    m_Detalles = detalles
End Sub