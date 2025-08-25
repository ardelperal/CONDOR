VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAuthRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================
' Clase: CAuthRepository
' Propósito: Implementación concreta del repositorio de autenticación
' Autor: CONDOR-Expert
' Fecha: 2025-01-15
' =====================================================

Option Explicit
Implements IAuthRepository

' Dependencias inyectadas
Private m_configService As IConfig
Private m_operationLogger As IOperationLogger
Private m_isInitialized As Boolean

' Inicializa el repositorio con sus dependencias
' @param configService: Servicio de configuración para obtener ruta del Backend
' @param operationLogger: Logger para registrar operaciones
Public Sub Initialize(ByVal configService As IConfig, ByVal operationLogger As IOperationLogger)
    If configService Is Nothing Then
        Err.Raise 5, "CAuthRepository.Initialize", "configService no puede ser Nothing"
    End If
    
    If operationLogger Is Nothing Then
        Err.Raise 5, "CAuthRepository.Initialize", "operationLogger no puede ser Nothing"
    End If
    
    Set m_configService = configService
    Set m_operationLogger = operationLogger
    m_isInitialized = True
End Sub

' Implementación de IAuthRepository_GetUserByEmail
Private Function IAuthRepository_GetUserByEmail(ByVal UserEmail As String) As DAO.Recordset
    Dim errorHandler As IErrorHandlerService
    
    If Not m_isInitialized Then
        Err.Raise 5, "CAuthRepository.GetUserByEmail", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, False, False, ";PWD=" & backendPassword)
    
    ' Consulta parametrizada para obtener datos de usuario por email
    sql = "SELECT EsAdministrador FROM TbUsuariosAplicaciones WHERE CorreoUsuario = ?"
    
    Set qdf = db.CreateQueryDef("", sql)
    qdf.Parameters(0) = UserEmail
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Cerrar conexión pero mantener recordset abierto
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Logging de auditoría
    m_operationLogger.LogOperation "READ", UserEmail, "Usuario consultado por email"
    
    Set IAuthRepository_GetUserByEmail = rs
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CAuthRepository.GetUserByEmail"
    m_operationLogger.LogOperation "READ_ERROR", UserEmail, "Error: " & Err.Description
    Set IAuthRepository_GetUserByEmail = Nothing
End Function

' Método público para acceso directo
Public Function GetUserByEmail(ByVal UserEmail As String) As DAO.Recordset
    Set GetUserByEmail = IAuthRepository_GetUserByEmail(UserEmail)
End Function

' Implementación de IAuthRepository_GetUserApplicationPermissions
Private Function IAuthRepository_GetUserApplicationPermissions(ByVal UserEmail As String, ByVal ApplicationID As String) As DAO.Recordset
    Dim errorHandler As IErrorHandlerService
    
    If Not m_isInitialized Then
        Err.Raise 5, "CAuthRepository.GetUserApplicationPermissions", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, False, False, ";PWD=" & backendPassword)
    
    ' Consulta parametrizada para obtener permisos específicos de aplicación
    sql = "SELECT EsUsuarioAdministrador, EsUsuarioCalidad, EsUsuarioTecnico " & _
          "FROM TbUsuariosAplicacionesPermisos " & _
          "WHERE CorreoUsuario = ? AND IDAplicacion = ?"
    
    Set qdf = db.CreateQueryDef("", sql)
    qdf.Parameters(0) = UserEmail
    qdf.Parameters(1) = ApplicationID
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Cerrar conexión pero mantener recordset abierto
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Logging de auditoría
    m_operationLogger.LogOperation "READ", UserEmail & "|" & ApplicationID, "Permisos de aplicación consultados"
    
    Set IAuthRepository_GetUserApplicationPermissions = rs
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CAuthRepository.GetUserApplicationPermissions"
    m_operationLogger.LogOperation "READ_ERROR", UserEmail & "|" & ApplicationID, "Error: " & Err.Description
    Set IAuthRepository_GetUserApplicationPermissions = Nothing
End Function

' Método público para obtener permisos específicos de aplicación
Public Function GetUserApplicationPermissions(ByVal UserEmail As String, ByVal ApplicationID As String) As DAO.Recordset
    Set GetUserApplicationPermissions = IAuthRepository_GetUserApplicationPermissions(UserEmail, ApplicationID)
End Function