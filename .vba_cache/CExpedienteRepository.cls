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
Private m_ErrorHandler As IErrorHandlerService

' ============================================================================
' MÉTODOS DE INICIALIZACIÓN
' ============================================================================

' Inicializa el repositorio con sus dependencias
Public Sub Initialize(ByVal config As IConfig, ByVal errorHandler As IErrorHandlerService)
    Set m_Config = config
    Set m_ErrorHandler = errorHandler
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IExpedienteRepository
' ============================================================================

' Implementación de IExpedienteRepository.ObtenerExpedientePorId
Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal params As Collection) As DAO.recordset
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim prm As QueryParameter
    
    ' Logging removido - responsabilidad de la capa de servicio
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_EXPEDIENTE_BY_ID)
    
    ' Asignar parámetros desde la colección
    For Each prm In params
        qdf.parameters(prm.ParameterName) = prm.ParameterValue
    Next prm
    
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
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientePorId"
    End If
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientePorId = Nothing
End Function

' Implementación de IExpedienteRepository.ObtenerExpedientePorNemotecnico
Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal params As Collection) As DAO.recordset
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim prm As QueryParameter
    
    ' Logging removido - responsabilidad de la capa de servicio
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_EXPEDIENTE_BY_NEMOTECNICO)
    
    ' Asignar parámetros desde la colección
    For Each prm In params
        qdf.parameters(prm.ParameterName) = prm.ParameterValue
    Next prm
    
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
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientePorNemotecnico"
    End If
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = Nothing
End Function

' Implementación de IExpedienteRepository.ObtenerExpedientesActivosParaSelector
Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector(ByVal params As Collection) As DAO.recordset
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim prm As QueryParameter
    
    ' Logging removido - responsabilidad de la capa de servicio
    
    ' Obtener conexión al Backend usando configuración
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False, ";PWD=" & m_Config.GetDatabasePassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", GET_EXPEDIENTES_ACTIVOS_SELECTOR)
    
    ' Asignar parámetros desde la colección (si los hay)
    If Not params Is Nothing Then
        For Each prm In params
            qdf.parameters(prm.ParameterName) = prm.ParameterValue
        Next prm
    End If
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset()
    
    ' Limpiar recursos de consulta (mantener recordset abierto)
    Set qdf = Nothing
    
    ' Devolver el recordset
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = rs
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then Set db = Nothing
    
    ' Log del error
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CExpedienteRepository.ObtenerExpedientesActivosParaSelector"
    End If
    
    ' Devolver Nothing en caso de error
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = Nothing
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' Método público de conveniencia para ObtenerExpedientePorId
Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As DAO.recordset
    Dim params As Collection
    Set params = New Collection
    params.Add CreateQueryParameter("pIdExpediente", idExpediente)
    Set ObtenerExpedientePorId = IExpedienteRepository_ObtenerExpedientePorId(params)
End Function

' Método público de conveniencia para ObtenerExpedientePorNemotecnico
Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As DAO.recordset
    Dim params As Collection
    Set params = New Collection
    params.Add CreateQueryParameter("pNemotecnico", Nemotecnico)
    Set ObtenerExpedientePorNemotecnico = IExpedienteRepository_ObtenerExpedientePorNemotecnico(params)
End Function

' Método público de conveniencia para ObtenerExpedientesActivosParaSelector
Public Function ObtenerExpedientesActivosParaSelector() As DAO.recordset
    Dim params As Collection
    Set params = New Collection
    ' Esta consulta no requiere parámetros, pero mantenemos la consistencia
    Set ObtenerExpedientesActivosParaSelector = IExpedienteRepository_ObtenerExpedientesActivosParaSelector(params)
End Function


