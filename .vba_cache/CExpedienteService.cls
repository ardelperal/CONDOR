Option Compare Database
Option Explicit

Implements IExpedienteService

' Dependencias inyectadas
Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger
Private m_SolicitudRepository As ISolicitudRepository

' ============================================================================
' MÉTODOS DE INICIALIZACIÓN
' ============================================================================

' Inicializa el servicio con sus dependencias
Public Sub Initialize(ByVal config As IConfig, _
                     ByVal operationLogger As IOperationLogger, _
                     ByVal solicitudRepository As ISolicitudRepository)
    Set m_Config = config
    Set m_OperationLogger = operationLogger
    Set m_SolicitudRepository = solicitudRepository
End Sub

' Implementación de la interfaz IExpedienteService
' Obtiene un expediente por su ID desde la base de datos de expedientes (método principal)
Public Function GetExpedienteById(ByVal idExpediente As Long) As T_Expediente
    Set GetExpedienteById = IExpedienteService_GetExpedienteById(idExpediente)
End Function

' Obtiene un expediente por su nemotécnico desde la base de datos de expedientes
Public Function GetExpedienteByNemotecnico(ByVal nemotecnico As String) As T_Expediente
    GetExpedienteByNemotecnico = IExpedienteService_GetExpedienteByNemotecnico(nemotecnico)
End Function

' Implementación de la interfaz para GetExpedienteById
Private Function IExpedienteService_GetExpedienteById(ByVal idExpediente As Long) As T_Expediente
    On Error GoTo ErrorHandler
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ExpedienteService.GetExpedienteById", "Consultando expediente ID: " & idExpediente
    End If
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim expediente As T_Expediente
    
    ' Obtener la base de datos actual
    Set db = CurrentDb
    
    ' Construir la consulta SQL parametrizada
    sql = "SELECT e.idExpediente, e.NumeroExpediente, e.Titulo, e.Descripcion, " & _
          "e.FechaCreacion, e.Estado, e.IdUsuarioCreador, " & _
          "u.NombreCompleto as NombreUsuarioCreador " & _
          "FROM tbExpedientes e " & _
          "LEFT JOIN tbUsuarios u ON e.IdUsuarioCreador = u.idUsuario " & _
          "WHERE e.idExpediente = ?"
    
    ' Crear QueryDef con parámetro
    Set qdf = db.CreateQueryDef("", sql)
    qdf.Parameters(0) = idExpediente
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Verificar si se encontró el expediente
    If Not rs.EOF Then
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
        ' Expediente no encontrado - devolver estructura vacía
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
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    
    ' Devolver el expediente
     IExpedienteService_GetExpedienteById = expediente
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then Set db = Nothing
    
    ' Log del error
    modErrorHandler.LogError "CExpedienteService.GetExpedienteById", Err.Number, Err.Description, "CExpedienteService.cls"
    
    ' Devolver estructura vacía en caso de error
     Dim expedienteVacio As T_Expediente
     expedienteVacio.idExpediente = 0
     IExpedienteService_GetExpedienteById = expedienteVacio
End Function



' Implementa IExpedienteService.GetExpedientesParaSelector
Private Function IExpedienteService_GetExpedientesParaSelector() As Object
    On Error GoTo ErrorHandler
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ExpedienteService.GetExpedientesParaSelector", "Obteniendo lista de expedientes para selector"
    End If
    
    ' Variables para la consulta
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Construir la consulta SQL para obtener expedientes activos
    sql = "SELECT e.idExpediente, e.NumeroExpediente, e.Titulo, e.Estado " & _
          "FROM tbExpedientes e " & _
          "WHERE e.Estado IN ('Activo', 'En Proceso') " & _
          "ORDER BY e.NumeroExpediente"
    
    ' Ejecutar consulta usando el repositorio inyectado
    Set rs = m_SolicitudRepository.ExecuteQuery(sql)
    
    ' Devolver el recordset
    Set IExpedienteService_GetExpedientesParaSelector = rs
    
    Exit Function
    
ErrorHandler:
    ' Log del error
    modErrorHandler.LogError "CExpedienteService.GetExpedientesParaSelector", Err.Number, Err.Description, "CExpedienteService.cls"
    
    ' Devolver Nothing en caso de error
    Set IExpedienteService_GetExpedientesParaSelector = Nothing
End Function

' Implementación de la interfaz para GetExpedienteByNemotecnico
Private Function IExpedienteService_GetExpedienteByNemotecnico(ByVal nemotecnico As String) As T_Expediente
    On Error GoTo ErrorHandler
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ExpedienteService.GetExpedienteByNemotecnico", "Consultando expediente por nemotécnico: " & nemotecnico, ""
    End If
    
    ' Usar el repositorio para obtener el recordset
    Dim rs As DAO.Recordset
    Set rs = m_SolicitudRepository.ObtenerExpedientePorNemotecnico(nemotecnico)
    
    Dim expediente As T_Expediente
    
    ' Verificar si se encontró el expediente
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
        ' Expediente no encontrado - devolver estructura vacía
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
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    ' Devolver el expediente
    IExpedienteService_GetExpedienteByNemotecnico = expediente
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    ' Log del error
    modErrorHandler.LogError Err.Number, Err.Description, "CExpedienteService.IExpedienteService_GetExpedienteByNemotecnico", "Nemotécnico: " & nemotecnico
    
    ' Devolver estructura vacía en caso de error
    Dim expedienteVacio As T_Expediente
    expedienteVacio.idExpediente = 0
    IExpedienteService_GetExpedienteByNemotecnico = expedienteVacio
End Function























