Option Compare Database
Option Explicit
Implements ISolicitud

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
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitud.Class_Initialize"
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
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitud.Initialize"
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
    ' Devuelve el objeto de datos que no sea nulo según el tipo de solicitud
    If m_tipoSolicitud = "PC" And Not m_datosPC Is Nothing Then
        Set ISolicitud_Datos = m_datosPC
    ElseIf m_tipoSolicitud = "CDCA" And Not m_datosCDCA Is Nothing Then
        Set ISolicitud_Datos = m_datosCDCA
    ElseIf m_tipoSolicitud = "CDCASUB" And Not m_datosCDCASUB Is Nothing Then
        Set ISolicitud_Datos = m_datosCDCASUB
    Else
        Set ISolicitud_Datos = Nothing
    End If
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
            ' Los datos específicos se cargarán según el tipo en el repositorio
            ' m_datosPC = loadedSolicitud.datosPC
            ISolicitud_Load = True
        Else
            ISolicitud_Load = False
        End If
    Else
        ISolicitud_Load = False
    End If
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitud.Load"
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
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitud.Save"
    ISolicitud_Save = False
End Function

Private Function ISolicitud_ChangeState(ByVal newState As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' TODO: Implementar lógica de cambio de estado
    ISolicitud_ChangeState = True
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitud.ChangeState"
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

Public Property Get subTipoSolicitud() As String
    subTipoSolicitud = m_subTipoSolicitud
End Property

Public Property Let subTipoSolicitud(ByVal val As String)
    m_subTipoSolicitud = val
End Property

Public Property Get fechaCreacion() As Date
    fechaCreacion = m_fechaCreacion
End Property

Public Property Let fechaCreacion(ByVal val As Date)
    m_fechaCreacion = val
End Property

Public Property Get usuarioCreacion() As String
    usuarioCreacion = m_usuarioCreacion
End Property

Public Property Let usuarioCreacion(ByVal val As String)
    m_usuarioCreacion = val
End Property

Public Property Get fechaPaseTecnico() As Date
    fechaPaseTecnico = m_fechaPaseTecnico
End Property

Public Property Let fechaPaseTecnico(ByVal val As Date)
    m_fechaPaseTecnico = val
End Property

Public Property Get fechaCompletadoTecnico() As Date
    fechaCompletadoTecnico = m_fechaCompletadoTecnico
End Property

Public Property Let fechaCompletadoTecnico(ByVal val As Date)
    m_fechaCompletadoTecnico = val
End Property

' --- Propiedades para objetos de datos específicos ---
Public Property Get datosPC() As T_Datos_PC
    Set datosPC = m_datosPC
End Property

Public Property Set datosPC(ByRef val As T_Datos_PC)
    Set m_datosPC = val
End Property

Public Property Get datosCDCA() As T_Datos_CD_CA
    Set datosCDCA = m_datosCDCA
End Property

Public Property Set datosCDCA(ByRef val As T_Datos_CD_CA)
    Set m_datosCDCA = val
End Property

Public Property Get datosCDCASUB() As T_Datos_CD_CA_SUB
    Set datosCDCASUB = m_datosCDCASUB
End Property

Public Property Set datosCDCASUB(ByRef val As T_Datos_CD_CA_SUB)
    Set m_datosCDCASUB = val
End Property

' --- Propiedad genérica Datos que devuelve el objeto activo ---
Public Property Get Datos() As Object
    Set Datos = ISolicitud_Datos
End Property












