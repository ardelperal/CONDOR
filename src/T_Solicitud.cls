Option Compare Database
Option Explicit
' Clase de tipo de datos para tbSolicitudes
' Basada en el Anexo A del CONDOR_MASTER_PLAN.md

Private m_idSolicitud As Long
Private m_idExpediente As String
Private m_tipoSolicitud As String
Private m_subTipoSolicitud As String
Private m_codigoSolicitud As String
Private m_estadoInterno As String
Private m_fechaCreacion As Date
Private m_usuarioCreacion As String
Private m_fechaPaseTecnico As Date
Private m_fechaCompletadoTecnico As Date

' Propiedades para datos específicos según tipo de solicitud
Private m_datosPC As T_Datos_PC
Private m_datosCDCA As T_Datos_CD_CA
Private m_datosCDCASUB As T_Datos_CD_CA_SUB

Public Property Get idSolicitud() As Long
    idSolicitud = m_idSolicitud
End Property

Public Property Let idSolicitud(ByVal value As Long)
    m_idSolicitud = value
End Property

Public Property Get idExpediente() As String
    idExpediente = m_idExpediente
End Property

Public Property Let idExpediente(ByVal value As String)
    m_idExpediente = value
End Property

Public Property Get tipoSolicitud() As String
    tipoSolicitud = m_tipoSolicitud
End Property

Public Property Let tipoSolicitud(ByVal value As String)
    m_tipoSolicitud = value
End Property

Public Property Get subTipoSolicitud() As String
    subTipoSolicitud = m_subTipoSolicitud
End Property

Public Property Let subTipoSolicitud(ByVal value As String)
    m_subTipoSolicitud = value
End Property

Public Property Get codigoSolicitud() As String
    codigoSolicitud = m_codigoSolicitud
End Property

Public Property Let codigoSolicitud(ByVal value As String)
    m_codigoSolicitud = value
End Property

Public Property Get estadoInterno() As String
    estadoInterno = m_estadoInterno
End Property

Public Property Let estadoInterno(ByVal value As String)
    m_estadoInterno = value
End Property

Public Property Get fechaCreacion() As Date
    fechaCreacion = m_fechaCreacion
End Property

Public Property Let fechaCreacion(ByVal value As Date)
    m_fechaCreacion = value
End Property

Public Property Get usuarioCreacion() As String
    usuarioCreacion = m_usuarioCreacion
End Property

Public Property Let usuarioCreacion(ByVal value As String)
    m_usuarioCreacion = value
End Property

Public Property Get fechaPaseTecnico() As Date
    fechaPaseTecnico = m_fechaPaseTecnico
End Property

Public Property Let fechaPaseTecnico(ByVal value As Date)
    m_fechaPaseTecnico = value
End Property

Public Property Get fechaCompletadoTecnico() As Date
    fechaCompletadoTecnico = m_fechaCompletadoTecnico
End Property

Public Property Let fechaCompletadoTecnico(ByVal value As Date)
    m_fechaCompletadoTecnico = value
End Property

' Propiedades para objetos de datos específicos
Public Property Get datosPC() As T_Datos_PC
    Set datosPC = m_datosPC
End Property

Public Property Set datosPC(ByVal value As T_Datos_PC)
    Set m_datosPC = value
End Property

Public Property Get datosCDCA() As T_Datos_CD_CA
    Set datosCDCA = m_datosCDCA
End Property

Public Property Set datosCDCA(ByVal value As T_Datos_CD_CA)
    Set m_datosCDCA = value
End Property

Public Property Get datosCDCASUB() As T_Datos_CD_CA_SUB
    Set datosCDCASUB = m_datosCDCASUB
End Property

Public Property Set datosCDCASUB(ByVal value As T_Datos_CD_CA_SUB)
    Set m_datosCDCASUB = value
End Property






