VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IAuthService

Private configSvc As IConfig
Private m_OperationLogger As IOperationLogger
Private m_AuthRepository As IAuthRepository

Private Sub Class_Initialize()
    ' InicializaciÃ³n sin dependencias
End Sub

Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, authRepository As IAuthRepository)
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Set configSvc = config
    Set m_OperationLogger = operationLogger
    Set m_AuthRepository = authRepository
    
    Exit Sub
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CAuthService.Initialize"
End Sub

Private Function IAuthService_GetUserRole(ByVal UserEmail As String) As E_UserRole
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler

    ' Logging de auditorÃ­a: Consulta de rol de usuario
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_QUERY", "Consultando rol para usuario: " & UserEmail, "CAuthService.GetUserRole"
    End If

    Dim role As E_UserRole
    Dim rs As DAO.Recordset

    role = Rol_Desconocido

    ' Primera consulta: verificar si es administrador global
    Set rs = m_AuthRepository.GetUserByEmail(UserEmail)

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            If rs!EsAdministrador = "SÃ­" Then
                role = Rol_Admin
            End If
        End If
        rs.Close
    End If
    
    Set rs = Nothing

    If role = Rol_Admin Then
        IAuthService_GetUserRole = Rol_Admin
        Exit Function
    End If

    ' Segunda consulta: roles específicos de la aplicación
    Set rs = m_AuthRepository.GetUserApplicationPermissions(UserEmail, configSvc.GetValue("IDAplicacion_CONDOR"))

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            If rs!EsUsuarioAdministrador = "SÃ­" Then
                role = Rol_Admin
            ElseIf rs!EsUsuarioCalidad = "SÃ­" Then
                role = Rol_Calidad
            ElseIf rs!EsUsuarioTecnico = "SÃ­" Then
                role = Rol_Tecnico
            End If
        End If
        rs.Close
    End If
    
    Set rs = Nothing

    ' Logging de auditorÃ­a: Resultado de consulta de rol
    If Not m_OperationLogger Is Nothing Then
        Dim roleText As String
        Select Case role
            Case Rol_Admin: roleText = "Administrador"
            Case Rol_Calidad: roleText = "Calidad"
            Case Rol_Tecnico: roleText = "TÃ©cnico"
            Case Else: roleText = "Desconocido"
        End Select
        m_OperationLogger.LogOperation "ROLE_ASSIGNED", "Rol asignado a " & UserEmail & ": " & roleText, "CAuthService.GetUserRole"
    End If
    
    IAuthService_GetUserRole = role
    Exit Function

ErrorHandler:
    IAuthService_GetUserRole = Rol_Desconocido
    If Not rs Is Nothing Then Set rs = Nothing
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_ERROR", "Error consultando rol para " & UserEmail & ": " & Err.Description, "CAuthService.GetUserRole"
    End If
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CAuthService.GetUserRole"
End Function

' MÃ©todo pÃºblico para acceso directo
Public Function GetUserRole(ByVal UserEmail As String) As E_UserRole
    GetUserRole = IAuthService_GetUserRole(UserEmail)
End Function











