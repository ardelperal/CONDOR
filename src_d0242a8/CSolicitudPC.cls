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

Implements ISolicitud

Private m_IdSolicitud As Long
Private m_IdExpediente As String
Private m_TipoSolicitud As String
Private m_CodigoSolicitud As String
Private m_EstadoInterno As String

Private m_DatosPC As T_Datos_PC
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

' --- 1. Implementaci├│n PRIVADA de la Interfaz ---
'     Satisface el contrato "Implements".
Private Property Get ISolicitud_idSolicitud() As Long
    ISolicitud_idSolicitud = m_IdSolicitud
End Property

Private Property Let ISolicitud_idSolicitud(ByVal RHS As Long)
    m_IdSolicitud = RHS
End Property

Private Property Get ISolicitud_IDExpediente() As String
    ISolicitud_IDExpediente = m_IdExpediente
End Property

Private Property Let ISolicitud_IDExpediente(ByVal RHS As String)
    m_IdExpediente = RHS
End Property

' Interfaz ISolicitud - Propiedades
Private Property Get ISolicitud_TipoSolicitud() As String
    ISolicitud_TipoSolicitud = m_TipoSolicitud
End Property

Private Property Let ISolicitud_TipoSolicitud(ByVal RHS As String)
    m_TipoSolicitud = RHS
End Property

Private Property Get ISolicitud_CodigoSolicitud() As String
    ISolicitud_CodigoSolicitud = m_CodigoSolicitud
End Property

Private Property Let ISolicitud_CodigoSolicitud(ByVal RHS As String)
    m_CodigoSolicitud = RHS
End Property

Private Property Get ISolicitud_EstadoInterno() As String
    ISolicitud_EstadoInterno = m_EstadoInterno
End Property

Private Property Let ISolicitud_EstadoInterno(ByVal RHS As String)
    m_EstadoInterno = RHS
End Property

Private Function ISolicitud_Load(ByVal ID As Long) As Boolean
    If Not mRepository Is Nothing Then
        Dim loadedSolicitud As CSolicitudPC
        Set loadedSolicitud = mRepository.LoadPC(ID)
        If Not loadedSolicitud Is Nothing Then
            ' Copiar propiedades del objeto cargado
            m_IdSolicitud = loadedSolicitud.idSolicitud
            m_IdExpediente = loadedSolicitud.IDExpediente
            m_TipoSolicitud = loadedSolicitud.TipoSolicitud
            m_CodigoSolicitud = loadedSolicitud.CodigoSolicitud
            m_EstadoInterno = loadedSolicitud.EstadoInterno
            m_DatosPC = loadedSolicitud.DatosPC
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
            m_IdSolicitud = savedId
            ISolicitud_Save = True
        Else
            ISolicitud_Save = False
        End If
    Else
        ISolicitud_Save = False
    End If
End Function

Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean
    ' TODO: Implementar l├│gica de cambio de estado
    ISolicitud_ChangeState = True
End Function

' --- 2. Miembros P├ÜBLICOS de la Clase ---
'     Exponen la funcionalidad al mundo. Delegan en la implementaci├│n de la interfaz.

Public Property Get idSolicitud() As Long
    idSolicitud = ISolicitud_idSolicitud
End Property

Public Property Let idSolicitud(ByVal val As Long)
    ISolicitud_idSolicitud = val
End Property

Public Property Get IDExpediente() As String
    IDExpediente = ISolicitud_IDExpediente
End Property

Public Property Let IDExpediente(ByVal val As String)
    ISolicitud_IDExpediente = val
End Property

Public Property Get TipoSolicitud() As String
    TipoSolicitud = ISolicitud_TipoSolicitud
End Property

Public Property Let TipoSolicitud(ByVal val As String)
    ISolicitud_TipoSolicitud = val
End Property

Public Property Get CodigoSolicitud() As String
    CodigoSolicitud = ISolicitud_CodigoSolicitud
End Property

Public Property Let CodigoSolicitud(ByVal val As String)
    ISolicitud_CodigoSolicitud = val
End Property

Public Property Get EstadoInterno() As String
    EstadoInterno = ISolicitud_EstadoInterno
End Property

Public Property Let EstadoInterno(ByVal val As String)
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
            m_IdSolicitud = savedId
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

' --- Propiedades espec├¡ficas de PC ---
Public Property Get DatosPC() As T_Datos_PC
    Set DatosPC = m_DatosPC
End Property

Public Property Set DatosPC(ByRef val As T_Datos_PC)
    Set m_DatosPC = val
End Property





