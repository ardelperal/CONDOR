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
Private m_OperationLogger As IOperationLogger

'******************************************************************************
' EVENTOS DE CLASE
'******************************************************************************

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set mRepository = modRepositoryFactory.CreateSolicitudRepository
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudPC.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Set mRepository = Nothing
    Set m_OperationLogger = Nothing
End Sub

' Método de inicialización para inyectar dependencias
Public Sub Initialize(operationLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    
    Set m_OperationLogger = operationLogger
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudPC.Initialize"
End Sub

' --- 1. Implementación PRIVADA de la Interfaz ---
'     Satisface el contrato "Implements".
Private Property Get ISolicitud_idSolicitud() As Long
    ISolicitud_idSolicitud = m_idSolicitud
End Property

Private Property Let ISolicitud_idSolicitud(ByVal RHS As Long)
    m_idSolicitud = RHS
End Property

Private Property Get ISolicitud_idExpediente() As String
    ISolicitud_idExpediente = m_idExpediente
End Property

Private Property Let ISolicitud_idExpediente(ByVal RHS As String)
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

Private Property Get ISolicitud_Datos() As Object
    Set ISolicitud_Datos = m_datosPC
End Property

Private Function ISolicitud_Load(ByVal ID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    If Not mRepository Is Nothing Then
        Dim loadedSolicitud As Object
        Set loadedSolicitud = mRepository.Load(ID)
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
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudPC.Load"
    ISolicitud_Load = False
End Function

Private Function ISolicitud_Save() As Boolean
    On Error GoTo ErrorHandler
    
    If Not mRepository Is Nothing Then
        Dim savedId As Long
        savedId = mRepository.Save(Me)
        If savedId > 0 Then
            m_idSolicitud = savedId
            ISolicitud_Save = True
        Else
            ISolicitud_Save = False
        End If
    Else
        ISolicitud_Save = False
    End If
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudPC.Save"
    ISolicitud_Save = False
End Function

Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' TODO: Implementar lógica de cambio de estado
    ISolicitud_ChangeState = True
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudPC.ChangeState"
    ISolicitud_ChangeState = False
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
    idExpediente = ISolicitud_idExpediente
End Property

Public Property Let idExpediente(ByVal val As String)
    ISolicitud_idExpediente = val
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







' --- Propiedades específicas de PC ---
Public Property Get datosPC() As T_Datos_PC
    Set datosPC = m_datosPC
End Property

Public Property Set datosPC(ByRef val As T_Datos_PC)
    Set m_datosPC = val
End Property












