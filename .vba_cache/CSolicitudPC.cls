Option Compare Database
Option Explicit
Implements ISolicitud

Private m_idSolicitud As Long
Private m_idExpediente As String
Private m_tipoSolicitud As String
Private m_codigoSolicitud As String
Private m_estadoInterno As String

Private m_datosPC As T_Datos_PC
Private mRepository As ISolicitudRepository

'******************************************************************************
' EVENTOS DE CLASE
'******************************************************************************

Private Sub Class_Initialize()
    Set mRepository = modRepositoryFactory.CreateSolicitudRepository
End Sub

Private Sub Class_Terminate()
    Set mRepository = Nothing
End Sub

' --- 1. Implementación PRIVADA de la Interfaz ---
'     Satisface el contrato "Implements".
Private Property Get ISolicitud_idSolicitud() As Long
    ISolicitud_idSolicitud = m_idSolicitud
End Property

Private Property Let ISolicitud_idSolicitud(ByVal RHS As Long)
    m_idSolicitud = RHS
End Property

Private Property Get ISolicitud_IDExpediente() As String
    ISolicitud_IDExpediente = m_idExpediente
End Property

Private Property Let ISolicitud_IDExpediente(ByVal RHS As String)
    m_idExpediente = RHS
End Property

' Interfaz ISolicitud - Propiedades
Private Property Get ISolicitud_TipoSolicitud() As String
    ISolicitud_TipoSolicitud = m_tipoSolicitud
End Property

Private Property Let ISolicitud_TipoSolicitud(ByVal RHS As String)
    m_tipoSolicitud = RHS
End Property

Private Property Get ISolicitud_CodigoSolicitud() As String
    ISolicitud_CodigoSolicitud = m_codigoSolicitud
End Property

Private Property Let ISolicitud_CodigoSolicitud(ByVal RHS As String)
    m_codigoSolicitud = RHS
End Property

Private Property Get ISolicitud_EstadoInterno() As String
    ISolicitud_EstadoInterno = m_estadoInterno
End Property

Private Property Let ISolicitud_EstadoInterno(ByVal RHS As String)
    m_estadoInterno = RHS
End Property

Private Function ISolicitud_Load(ByVal ID As Long) As Boolean
    If Not mRepository Is Nothing Then
        Dim loadedSolicitud As CSolicitudPC
        Set loadedSolicitud = mRepository.LoadPC(ID)
        If Not loadedSolicitud Is Nothing Then
            ' Copiar propiedades del objeto cargado
            m_idSolicitud = loadedSolicitud.idSolicitud
            m_idExpediente = loadedSolicitud.idExpediente
            m_tipoSolicitud = loadedSolicitud.tipoSolicitud
            m_codigoSolicitud = loadedSolicitud.codigoSolicitud
            m_estadoInterno = loadedSolicitud.estadoInterno
            m_datosPC = loadedSolicitud.datosPC
            ISolicitud_Load = True
        Else
            ISolicitud_Load = False
        End If
    Else
        ISolicitud_Load = False
    End If
End Function

Private Function ISolicitud_Save() As Boolean
    If Not mRepository Is Nothing Then
        Dim savedId As Long
        savedId = mRepository.SavePC(Me)
        If savedId > 0 Then
            m_idSolicitud = savedId
            ISolicitud_Save = True
        Else
            ISolicitud_Save = False
        End If
    Else
        ISolicitud_Save = False
    End If
End Function

Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean
    ' TODO: Implementar lógica de cambio de estado
    ISolicitud_ChangeState = True
End Function

' --- 2. Miembros PÚBLICOS de la Clase ---
'     Exponen la funcionalidad al mundo. Delegan en la implementación de la interfaz.

Public Property Get idSolicitud() As Long
    idSolicitud = ISolicitud_idSolicitud
End Property

Public Property Let idSolicitud(ByVal val As Long)
    ISolicitud_idSolicitud = val
End Property

Public Property Get idExpediente() As String
    idExpediente = ISolicitud_IDExpediente
End Property

Public Property Let idExpediente(ByVal val As String)
    ISolicitud_IDExpediente = val
End Property

Public Property Get tipoSolicitud() As String
    tipoSolicitud = ISolicitud_TipoSolicitud
End Property

Public Property Let tipoSolicitud(ByVal val As String)
    ISolicitud_TipoSolicitud = val
End Property

Public Property Get codigoSolicitud() As String
    codigoSolicitud = ISolicitud_CodigoSolicitud
End Property

Public Property Let codigoSolicitud(ByVal val As String)
    ISolicitud_CodigoSolicitud = val
End Property

Public Property Get estadoInterno() As String
    estadoInterno = ISolicitud_EstadoInterno
End Property

Public Property Let estadoInterno(ByVal val As String)
    ISolicitud_EstadoInterno = val
End Property

Public Function Load(ByVal ID As Long) As Boolean
    Load = ISolicitud_Load(ID)
End Function

Public Function Save() As Boolean
    If Not mRepository Is Nothing Then
        Dim savedId As Long
        savedId = mRepository.SavePC(Me)
        If savedId > 0 Then
            m_idSolicitud = savedId
            Save = True
        Else
            Save = False
        End If
    Else
        Save = False
    End If
End Function

Public Function ChangeState(ByVal newState As String) As Boolean
    ChangeState = ISolicitud_ChangeState(newState)
End Function

' --- Propiedades específicas de PC ---
Public Property Get datosPC() As T_Datos_PC
    Set datosPC = m_datosPC
End Property

Public Property Set datosPC(ByRef val As T_Datos_PC)
    Set m_datosPC = val
End Property












