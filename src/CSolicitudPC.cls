VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolicitudPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ============================================================================
' Clase: CSolicitudPC
' Descripción: Implementación concreta para solicitudes de Propuesta de Cambio
' Implementa: ISolicitud
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

Implements ISolicitud

' Datos específicos de la Propuesta de Cambio
Private m_DatosPC As T_Datos_PC

' Variables privadas para las propiedades de la interfaz
Private m_ID_Solicitud As Long
Private m_ID_Expediente As String
Private m_TipoSolicitud As String
Private m_CodigoSolicitud As String
Private m_EstadoInterno As String

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ ISolicitud
' ============================================================================

' Propiedades de la interfaz ISolicitud
Private Property Get ISolicitud_ID_Solicitud() As Long
    ISolicitud_ID_Solicitud = m_ID_Solicitud
End Property

Private Property Let ISolicitud_ID_Solicitud(ByVal value As Long)
    m_ID_Solicitud = value
End Property

Private Property Get ISolicitud_ID_Expediente() As String
    ISolicitud_ID_Expediente = m_ID_Expediente
End Property

Private Property Let ISolicitud_ID_Expediente(ByVal value As String)
    m_ID_Expediente = value
End Property

Private Property Get ISolicitud_TipoSolicitud() As String
    ISolicitud_TipoSolicitud = m_TipoSolicitud
End Property

Private Property Let ISolicitud_TipoSolicitud(ByVal value As String)
    m_TipoSolicitud = value
End Property

Private Property Get ISolicitud_CodigoSolicitud() As String
    ISolicitud_CodigoSolicitud = m_CodigoSolicitud
End Property

Private Property Let ISolicitud_CodigoSolicitud(ByVal value As String)
    m_CodigoSolicitud = value
End Property

Private Property Get ISolicitud_EstadoInterno() As String
    ISolicitud_EstadoInterno = m_EstadoInterno
End Property

Private Property Let ISolicitud_EstadoInterno(ByVal value As String)
    m_EstadoInterno = value
End Property

' Métodos de la interfaz ISolicitud
Private Function ISolicitud_Load(ByVal ID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    
    ' Obtener datos usando modDatabase
    Set rs = modDatabase.GetSolicitudData(ID)
    
    ' Verificar si se encontraron datos
    If rs Is Nothing Or rs.EOF Then
        ISolicitud_Load = False
        Exit Function
    End If
    
    ' Hidratar propiedades de la clase con datos del Recordset
    With rs
        m_ID_Solicitud = !ID
        m_ID_Expediente = !NumeroExpediente & ""
        m_TipoSolicitud = !TipoSolicitud & ""
        m_EstadoInterno = !EstadoInterno & ""
        
        ' Generar código de solicitud
        m_CodigoSolicitud = m_TipoSolicitud & "-" & Format(m_ID_Solicitud, "0000")
        
        ' Hidratar estructura DatosPC
        m_DatosPC.ID = !DatosID
        m_DatosPC.SolicitudID = !ID
        m_DatosPC.NumeroExpediente = !NumeroExpediente & ""
        m_DatosPC.TipoSolicitud = !TipoSolicitud & ""
        m_DatosPC.DescripcionCambio = !DescripcionCambio & ""
        m_DatosPC.JustificacionCambio = !JustificacionCambio & ""
        m_DatosPC.ImpactoSeguridad = !ImpactoSeguridad & ""
        m_DatosPC.ImpactoCalidad = !ImpactoCalidad & ""
        m_DatosPC.FechaCreacion = !FechaCreacion
        m_DatosPC.FechaUltimaModificacion = !FechaUltimaModificacion
        m_DatosPC.Estado = !EstadoDatos & ""
        m_DatosPC.Activo = !Activo
    End With
    
    rs.Close
    Set rs = Nothing
    
    ISolicitud_Load = True
    Exit Function
    
ErrorHandler:
    ISolicitud_Load = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function ISolicitud_Save() As Boolean
    On Error GoTo ErrorHandler
    
    Dim solicitudData As T_Solicitud
    
    ' Crear estructura T_Solicitud con los datos de la clase
    With solicitudData
        .ID = m_ID_Solicitud
        .NumeroExpediente = m_ID_Expediente
        .TipoSolicitud = m_TipoSolicitud
        .EstadoInterno = m_EstadoInterno
        .EstadoRAC = "" ' Por defecto vacío
        .FechaCreacion = m_DatosPC.FechaCreacion
        .FechaUltimaModificacion = Now()
        .Usuario = Environ("USERNAME") ' Usuario actual del sistema
        .Observaciones = "" ' Por defecto vacío
        .Activo = True
    End With
    
    ' Sincronizar datos entre solicitud y DatosPC
    m_DatosPC.SolicitudID = m_ID_Solicitud
    m_DatosPC.NumeroExpediente = m_ID_Expediente
    m_DatosPC.TipoSolicitud = m_TipoSolicitud
    m_DatosPC.FechaUltimaModificacion = Now()
    m_DatosPC.Activo = True
    
    ' Llamar a modDatabase.SaveSolicitudPC
    ISolicitud_Save = modDatabase.SaveSolicitudPC(solicitudData, m_DatosPC)
    
    ' Actualizar ID si es un registro nuevo
    If ISolicitud_Save And m_ID_Solicitud = 0 Then
        m_ID_Solicitud = solicitudData.ID
        m_CodigoSolicitud = m_TipoSolicitud & "-" & Format(m_ID_Solicitud, "0000")
    End If
    
    Exit Function
    
ErrorHandler:
    ISolicitud_Save = False
End Function

Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean
    ' TODO: Implementar cambio de estado siguiendo las reglas del workflow
    ' Por ahora retorna True para que compile
    On Error GoTo ErrorHandler
    
    ' Aquí se implementará la lógica de cambio de estado
    m_EstadoInterno = newState
    
    ISolicitud_ChangeState = True
    Exit Function
    
ErrorHandler:
    ISolicitud_ChangeState = False
End Function

' ============================================================================
' PROPIEDADES PÚBLICAS ESPECÍFICAS DE ESTA CLASE
' ============================================================================

Public Property Get DatosPC() As T_Datos_PC
    DatosPC = m_DatosPC
End Property

Public Property Let DatosPC(ByVal value As T_Datos_PC)
    m_DatosPC = value
End Property

' ============================================================================
' INICIALIZACIÓN
' ============================================================================

Private Sub Class_Initialize()
    ' Inicializar valores por defecto
    m_TipoSolicitud = "PC"
    m_EstadoInterno = "BORRADOR"
End Sub


