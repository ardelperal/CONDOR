Option Compare Database
Option Explicit

' Versión 2.1 - Sincronizada con el esquema de la base de datos

Private m_FechaHora As Date
Private m_Usuario As String
Private m_TipoOperacion As String
Private m_Entidad As String
Private m_IDEntidadAfectada As Long
Private m_Descripcion As String
Private m_Resultado As String
Private m_Detalles As String

Public Property Get fechaHora() As Date: fechaHora = m_FechaHora: End Property
Public Property Let fechaHora(ByVal value As Date): m_FechaHora = value: End Property

Public Property Get usuario() As String: usuario = m_Usuario: End Property
Public Property Let usuario(ByVal value As String): m_Usuario = value: End Property

Public Property Get tipoOperacion() As String: tipoOperacion = m_TipoOperacion: End Property
Public Property Let tipoOperacion(ByVal value As String): m_TipoOperacion = value: End Property

Public Property Get entidad() As String: entidad = m_Entidad: End Property
Public Property Let entidad(ByVal value As String): m_Entidad = value: End Property

Public Property Get idEntidadAfectada() As Long: idEntidadAfectada = m_IDEntidadAfectada: End Property
Public Property Let idEntidadAfectada(ByVal value As Long): m_IDEntidadAfectada = value: End Property

Public Property Get descripcion() As String: descripcion = m_Descripcion: End Property
Public Property Let descripcion(ByVal value As String): m_Descripcion = value: End Property

Public Property Get resultado() As String: resultado = m_Resultado: End Property
Public Property Let resultado(ByVal value As String): m_Resultado = value: End Property

Public Property Get detalles() As String: detalles = m_Detalles: End Property
Public Property Let detalles(ByVal value As String): m_Detalles = value: End Property

Public Sub Initialize(ByVal fechaHora As Date, ByVal usuario As String, ByVal tipoOperacion As String, ByVal entidad As String, ByVal idEntidadAfectada As Long, ByVal descripcion As String, ByVal resultado As String, ByVal detalles As String)
    m_FechaHora = fechaHora
    m_Usuario = usuario
    m_TipoOperacion = tipoOperacion
    m_Entidad = entidad
    m_IDEntidadAfectada = idEntidadAfectada
    m_Descripcion = descripcion
    m_Resultado = resultado
    m_Detalles = detalles
End Sub