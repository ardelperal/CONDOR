VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpedienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IExpedienteRepository

' Dependencias inyectadas
Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger

' ============================================================================
' MÉTODOS DE INICIALIZACIÓN
' ============================================================================

' Inicializa el repositorio con sus dependencias
Public Sub Initialize(ByVal config As IConfig, _
                     ByVal operationLogger As IOperationLogger)
    Set m_Config = config
    Set m_OperationLogger = operationLogger
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IExpedienteRepository
' ============================================================================

' Implementación de IExpedienteRepository.ObtenerExpedientePorId
Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal IdExpediente As Long) As DAO.Recordset
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "READ", CStr(IdExpediente), "ExpedienteRepository: Consultando expediente ID: " & IdExpediente
    End If
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), False, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Construir la consulta SQL parametrizada
    sql = "SELECT e.idExpediente, e.NumeroExpediente, e.Titulo, e.Descripcion, " & _
          "e.FechaCreacion, e.Estado, e.IdUsuarioCreador, " & _
          "u.NombreCompleto as NombreUsuarioCreador " & _
          "FROM tbExpedientes e " & _
          "LEFT JOIN tbUsuarios u ON e.IdUsuarioCreador = u.idUsuario " & _
          "WHERE e.idExpediente = ?"
    
    ' Crear QueryDef con parámetro
    Set qdf = db.CreateQueryDef("", sql)
    qdf.Parameters(0) = IdExpediente
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Limpiar recursos de consulta (mantener recordset abierto)
    Set qdf = Nothing
    
    ' Devolver el recordset
    Set IExpedienteRepository_ObtenerExpedientePorId = rs
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then Set db = Nothing
    
    ' Log del error
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientePorId"
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientePorId = Nothing
End Function

' Implementación de IExpedienteRepository.ObtenerExpedientePorNemotecnico
Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As DAO.Recordset
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "READ", Nemotecnico, "ExpedienteRepository: Consultando expediente por nemotécnico: " & Nemotecnico
    End If
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), False, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Construir la consulta SQL parametrizada
    sql = "SELECT e.idExpediente, e.NumeroExpediente, e.Titulo, e.Descripcion, " & _
          "e.FechaCreacion, e.Estado, e.IdUsuarioCreador, " & _
          "u.NombreCompleto as NombreUsuarioCreador " & _
          "FROM tbExpedientes e " & _
          "LEFT JOIN tbUsuarios u ON e.IdUsuarioCreador = u.idUsuario " & _
          "WHERE e.NumeroExpediente = ?"
    
    ' Crear QueryDef con parámetro
    Set qdf = db.CreateQueryDef("", sql)
    qdf.Parameters(0) = Nemotecnico
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Limpiar recursos de consulta (mantener recordset abierto)
    Set qdf = Nothing
    
    ' Devolver el recordset
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = rs
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then Set db = Nothing
    
    ' Log del error
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientePorNemotecnico"
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = Nothing
End Function

' Implementación de IExpedienteRepository.ObtenerExpedientesActivosParaSelector
Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector() As DAO.Recordset
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Log de la operación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "READ", "", "ExpedienteRepository: Obteniendo expedientes activos para selector"
    End If
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), False, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Construir la consulta SQL para obtener expedientes activos
    sql = "SELECT e.idExpediente, e.NumeroExpediente, e.Titulo, e.Estado " & _
          "FROM tbExpedientes e " & _
          "WHERE e.Estado IN ('Activo', 'En Proceso') " & _
          "ORDER BY e.NumeroExpediente"
    
    ' Ejecutar consulta
    Set rs = db.OpenRecordset(sql)
    
    ' Devolver el recordset
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = rs
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then Set db = Nothing
    
    ' Log del error
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientesActivosParaSelector"
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = Nothing
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' Método público de conveniencia para ObtenerExpedientePorId
Public Function ObtenerExpedientePorId(ByVal IdExpediente As Long) As DAO.Recordset
    Set ObtenerExpedientePorId = IExpedienteRepository_ObtenerExpedientePorId(IdExpediente)
End Function

' Método público de conveniencia para ObtenerExpedientePorNemotecnico
Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As DAO.Recordset
    Set ObtenerExpedientePorNemotecnico = IExpedienteRepository_ObtenerExpedientePorNemotecnico(Nemotecnico)
End Function

' Método público de conveniencia para ObtenerExpedientesActivosParaSelector
Public Function ObtenerExpedientesActivosParaSelector() As DAO.Recordset
    Set ObtenerExpedientesActivosParaSelector = IExpedienteRepository_ObtenerExpedientesActivosParaSelector()
End Function