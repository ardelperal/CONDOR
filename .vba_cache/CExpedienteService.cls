Option Compare Database
Option Explicit

Implements IExpedienteService

' Dependencias inyectadas
Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger
Private m_ExpedienteRepository As IExpedienteRepository

' ============================================================================
' MÃ‰TODOS DE INICIALIZACIÃ“N
' ============================================================================

' Inicializa el servicio con sus dependencias
Public Sub Initialize(ByVal config As IConfig, _
                     ByVal operationLogger As IOperationLogger, _
                     ByVal expedienteRepository As IExpedienteRepository)
    Set m_Config = config
    Set m_OperationLogger = operationLogger
    Set m_ExpedienteRepository = expedienteRepository
End Sub

' ImplementaciÃ³n de la interfaz IExpedienteService
' Obtiene un expediente por su ID desde la base de datos de expedientes (mÃ©todo principal)
Public Function GetExpedienteById(ByVal idExpediente As Long) As T_Expediente
    Set GetExpedienteById = IExpedienteService_GetExpedienteById(idExpediente)
End Function

' Obtiene un expediente por su nemotécnico desde la base de datos de expedientes
Public Function GetExpedienteByNemotecnico(ByVal nemotecnico As String) As T_Expediente
    Set GetExpedienteByNemotecnico = IExpedienteService_GetExpedienteByNemotecnico(nemotecnico)
End Function

' ImplementaciÃ³n de la interfaz para GetExpedienteById
Private Function IExpedienteService_GetExpedienteById(ByVal idExpediente As Long) As T_Expediente
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables
    Dim rs As DAO.Recordset
    Dim expediente As T_Expediente
    
    ' Instanciar el objeto expediente
    Set expediente = New T_Expediente
    
    ' Log de la operaciÃ³n
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "READ", CStr(idExpediente), "ExpedienteService: Consultando expediente ID: " & idExpediente
    End If
    
    ' Usar el repositorio para obtener el recordset
    Set rs = m_ExpedienteRepository.ObtenerExpedientePorId(idExpediente)
    
    ' Verificar si se encontrÃ³ el expediente
    If Not rs Is Nothing And Not rs.EOF Then
        ' Mapear los datos del recordset a la estructura T_Expediente
        expediente.idExpediente = rs("idExpediente")
        expediente.NumeroExpediente = rs("NumeroExpediente") & ""
        expediente.Titulo = rs("Titulo") & ""
        expediente.Descripcion = rs("Descripcion") & ""
        expediente.FechaCreacion = rs("FechaCreacion")
        expediente.Estado = rs("Estado") & ""
        expediente.IdUsuarioCreador = rs("IdUsuarioCreador")
        expediente.NombreUsuarioCreador = rs("NombreUsuarioCreador") & ""
    Else
        ' Expediente no encontrado - devolver estructura vacÃ­a
        expediente.idExpediente = 0
        expediente.NumeroExpediente = ""
        expediente.Titulo = ""
        expediente.Descripcion = ""
        expediente.FechaCreacion = #1/1/1900#
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
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CExpedienteService.GetExpedienteById"
    
    ' Devolver estructura vacía en caso de error
    Dim expedienteVacio As T_Expediente
    Set expedienteVacio = New T_Expediente
    expedienteVacio.idExpediente = 0
    Set IExpedienteService_GetExpedienteById = expedienteVacio
End Function



' Implementa IExpedienteService.GetExpedientesParaSelector
Private Function IExpedienteService_GetExpedientesParaSelector() As Object
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta
    Dim rs As DAO.Recordset
    
    ' Log de la operaciÃ³n
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
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CExpedienteService.GetExpedientesParaSelector"
    
    ' Devolver Nothing en caso de error
    Set IExpedienteService_GetExpedientesParaSelector = Nothing
End Function

' ImplementaciÃ³n de la interfaz para GetExpedienteByNemotecnico
Private Function IExpedienteService_GetExpedienteByNemotecnico(ByVal nemotecnico As String) As T_Expediente
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables
    Dim rs As DAO.Recordset
    Dim expediente As T_Expediente
    
    ' Log de la operaciÃ³n
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ExpedienteService.GetExpedienteByNemotecnico", "Consultando expediente por nemotÃ©cnico: " & nemotecnico, ""
    End If
    
    ' Usar el repositorio para obtener el recordset
    Set rs = m_ExpedienteRepository.ObtenerExpedientePorNemotecnico(nemotecnico)
    
    ' Verificar si se encontrÃ³ el expediente
    If Not rs Is Nothing And Not rs.EOF Then
        ' Mapear los datos del recordset a la estructura T_Expediente
        expediente.idExpediente = rs("idExpediente")
        expediente.NumeroExpediente = rs("NumeroExpediente") & ""
        expediente.Titulo = rs("Titulo") & ""
        expediente.Descripcion = rs("Descripcion") & ""
        expediente.FechaCreacion = rs("FechaCreacion")
        expediente.Estado = rs("Estado") & ""
        expediente.IdUsuarioCreador = rs("IdUsuarioCreador")
        expediente.NombreUsuarioCreador = rs("NombreUsuarioCreador") & ""
    Else
        ' Expediente no encontrado - devolver estructura vacÃ­a
        expediente.idExpediente = 0
        expediente.NumeroExpediente = ""
        expediente.Titulo = ""
        expediente.Descripcion = ""
        expediente.FechaCreacion = #1/1/1900#
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
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CExpedienteService.IExpedienteService_GetExpedienteByNemotecnico", "NemotÃ©cnico: " & nemotecnico
    
    ' Devolver estructura vacía en caso de error
    Dim expedienteVacio As T_Expediente
    Set expedienteVacio = New T_Expediente
    expedienteVacio.idExpediente = 0
    Set IExpedienteService_GetExpedienteByNemotecnico = expedienteVacio
End Function























