VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Implements IExpedienteService

' Dependencias inyectadas
Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger
Private m_ExpedienteRepository As IExpedienteRepository
Private m_ErrorHandler As IErrorHandlerService

' ============================================================================
' MÃ‰TODOS DE INICIALIZACIÃ“N
' ============================================================================

' Inicializa el servicio con sus dependencias
Public Sub Initialize(ByVal config As IConfig, _
                     ByVal operationLogger As IOperationLogger, _
                     ByVal expedienteRepository As IExpedienteRepository, _
                     ByVal errorHandler As IErrorHandlerService)
    Set m_Config = config
    Set m_OperationLogger = operationLogger
    Set m_ExpedienteRepository = expedienteRepository
    Set m_ErrorHandler = errorHandler
End Sub

' Implementación de la interfaz IExpedienteService
' Obtiene un expediente por su ID desde la base de datos de expedientes (método principal)
Public Function GetExpedienteById(ByVal idExpediente As Long) As EExpediente
    Set GetExpedienteById = IExpedienteService_GetExpedienteById(idExpediente)
End Function

' Obtiene un expediente por su nemotécnico desde la base de datos de expedientes
Public Function GetExpedienteByNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set GetExpedienteByNemotecnico = IExpedienteService_GetExpedienteByNemotecnico(Nemotecnico)
End Function

' Obtiene expedientes para selector
Public Function GetExpedientesParaSelector() As DAO.Recordset
    Set GetExpedientesParaSelector = IExpedienteService_GetExpedientesParaSelector()
End Function

' Implementación de la interfaz para GetExpedienteById
Private Function IExpedienteService_GetExpedienteById(ByVal idExpediente As Long) As EExpediente
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables
    Dim rs As DAO.recordset
    Dim expediente As EExpediente
    
    ' Instanciar el objeto expediente
    Set expediente = New EExpediente
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "READ", CStr(idExpediente), "ExpedienteService: Consultando expediente ID: " & idExpediente
    End If
    
    ' Usar el repositorio para obtener el recordset
    Set rs = m_ExpedienteRepository.ObtenerExpedientePorId(idExpediente)
    
    ' Verificar si se encontró el expediente
    If Not rs Is Nothing And Not rs.EOF Then
        ' Mapear los datos del recordset a la estructura EExpediente
        expediente.idExpediente = rs("idExpediente")
        expediente.NumeroExpediente = rs("NumeroExpediente") & ""
        expediente.Titulo = rs("Titulo") & ""
        expediente.Descripcion = rs("Descripcion") & ""
        expediente.fechaCreacion = rs("FechaCreacion")
        expediente.Estado = rs("Estado") & ""
        expediente.IdUsuarioCreador = rs("IdUsuarioCreador")
        expediente.NombreUsuarioCreador = rs("NombreUsuarioCreador") & ""
    Else
        ' Expediente no encontrado - devolver estructura vacía
        expediente.idExpediente = 0
        expediente.NumeroExpediente = ""
        expediente.Titulo = ""
        expediente.Descripcion = ""
        expediente.fechaCreacion = #1/1/1900#
        expediente.Estado = ""
        expediente.IdUsuarioCreador = 0
        expediente.NombreUsuarioCreador = ""
    End If
    
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    ' Devolver el expediente
    Set IExpedienteService_GetExpedienteById = expediente
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    ' Log del error
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteService.GetExpedienteById"
    End If
    
    ' Devolver estructura vacía en caso de error
    Dim expedienteVacio As EExpediente
    Set expedienteVacio = New EExpediente
    expedienteVacio.idExpediente = 0
    Set IExpedienteService_GetExpedienteById = expedienteVacio
End Function



' Implementa IExpedienteService.GetExpedientesParaSelector
Private Function IExpedienteService_GetExpedientesParaSelector() As DAO.Recordset
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta
    Dim rs As DAO.recordset
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "READ", "", "ExpedienteService: Obteniendo lista de expedientes para selector"
    End If
    
    ' Usar el repositorio para obtener expedientes activos
    Set rs = m_ExpedienteRepository.ObtenerExpedientesActivosParaSelector()
    
    ' Devolver el recordset
    Set IExpedienteService_GetExpedientesParaSelector = rs
    
    Exit Function
    
ErrorHandler:
    ' Log del error
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteService.GetExpedientesParaSelector"
    End If
    
    ' Devolver Nothing en caso de error
    Set IExpedienteService_GetExpedientesParaSelector = Nothing
End Function

' Implementación de la interfaz para GetExpedienteByNemotecnico
Private Function IExpedienteService_GetExpedienteByNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables
    Dim rs As DAO.recordset
    Dim expediente As EExpediente
    
    ' Instanciar el objeto expediente
    Set expediente = New EExpediente
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ExpedienteService.GetExpedienteByNemotecnico", "Consultando expediente por nemotécnico: " & Nemotecnico, ""
    End If
    
    ' Usar el repositorio para obtener el recordset
    Set rs = m_ExpedienteRepository.ObtenerExpedientePorNemotecnico(Nemotecnico)
    
    ' Verificar si se encontró el expediente
    If Not rs Is Nothing And Not rs.EOF Then
        ' Mapear los datos del recordset a la estructura EExpediente
        expediente.idExpediente = rs("idExpediente")
        expediente.NumeroExpediente = rs("NumeroExpediente") & ""
        expediente.Titulo = rs("Titulo") & ""
        expediente.Descripcion = rs("Descripcion") & ""
        expediente.fechaCreacion = rs("FechaCreacion")
        expediente.Estado = rs("Estado") & ""
        expediente.IdUsuarioCreador = rs("IdUsuarioCreador")
        expediente.NombreUsuarioCreador = rs("NombreUsuarioCreador") & ""
    Else
        ' Expediente no encontrado - devolver estructura vacía
        expediente.idExpediente = 0
        expediente.NumeroExpediente = ""
        expediente.Titulo = ""
        expediente.Descripcion = ""
        expediente.fechaCreacion = #1/1/1900#
        expediente.Estado = ""
        expediente.IdUsuarioCreador = 0
        expediente.NombreUsuarioCreador = ""
    End If
    
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    ' Devolver el expediente
    Set IExpedienteService_GetExpedienteByNemotecnico = expediente
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    ' Log del error
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteService.IExpedienteService_GetExpedienteByNemotecnico"
    End If
    
    ' Devolver estructura vacía en caso de error
    Dim expedienteVacio As EExpediente
    Set expedienteVacio = New EExpediente
    expedienteVacio.idExpediente = 0
    Set IExpedienteService_GetExpedienteByNemotecnico = expedienteVacio
End Function


























