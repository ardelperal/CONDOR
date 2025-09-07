VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COperationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' Clase: COperationRepository
' Descripción: Implementación concreta del repositorio de logs de operaciones.
'              Maneja la persistencia en la tabla Tb_Operaciones_Log.
' Arquitectura: Capa de Datos - Implementación del patrón Repository

Implements IOperationRepository

' Dependencias inyectadas
Private m_configService As IConfig
Private m_ErrorHandler As IErrorHandlerService

' Método de inicialización para inyectar dependencias
Public Sub Initialize(config As IConfig, errorHandler As IErrorHandlerService)
    On Error GoTo errorHandler
    Set m_configService = config
    Set m_ErrorHandler = errorHandler
    Exit Sub
errorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationRepository.Initialize"
End Sub

' Implementación del contrato IOperationRepository
Private Sub IOperationRepository_SaveLog(ByVal operationType As String, ByVal entityId As Long, ByVal details As String)
    On Error GoTo errorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim logDbPath As String
    Dim logDbPassword As String
    Dim sql As String
    
    ' Obtener la ruta y contraseña de la base de datos de logs de operaciones
    logDbPath = m_configService.GetCondorDataPath()
    logDbPassword = m_configService.GetCondorPassword()
    
    Set db = DBEngine.OpenDatabase(logDbPath, dbFailOnError, False, ";PWD=" & logDbPassword)
    
    ' Usar la consulta centralizada con parámetros con nombre
    Set qdf = db.CreateQueryDef("", INSERT_OPERATION_LOG)
    qdf.Parameters("pFechaHora") = Now
    qdf.Parameters("pUsuario") = Environ("USERNAME") ' O el usuario autenticado de la aplicación
    qdf.Parameters("pTipoOperacion") = operationType
    qdf.Parameters("pIDEntidadAfectada") = entityId
    qdf.Parameters("pDetalles") = details
    
    ' Ejecutar la consulta parametrizada
    qdf.Execute
    
    ' Limpiar recursos
    Set qdf = Nothing
    db.Close
    
    Exit Sub
    
errorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationRepository.IOperationRepository_SaveLog", _
                                  "TipoOperacion: " & operationType & ", Entidad: " & entityId
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Sub SaveLog(ByVal operationType As String, ByVal entityId As Long, ByVal details As String)
    IOperationRepository_SaveLog operationType, entityId, details
End Sub





