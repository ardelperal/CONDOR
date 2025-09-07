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
Private m_ErrorHandler As IErrorHandlerService

Private Sub Class_Initialize()
    ' Inicialización sin dependencias
End Sub

Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, authRepository As IAuthRepository, errorHandler As IErrorHandlerService)
    On Error GoTo errorHandler
    
    Set configSvc = config
    Set m_OperationLogger = operationLogger
    Set m_AuthRepository = authRepository
    Set m_ErrorHandler = errorHandler
    
    Exit Sub
    
errorHandler:
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthService.Initialize"
    End If
End Sub

Private Function IAuthService_GetUserRole(ByVal userEmail As String) As UserRole
    On Error GoTo errorHandler

    ' Logging de auditoría: Consulta de rol de usuario
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_QUERY", 0, "Consultando rol para usuario: " & userEmail & " (CAuthService.GetUserRole)"
    End If

    Dim role As UserRole
    Dim authData As EAuthData

    role = RolDesconocido

    ' Obtener todos los datos de autenticación en una sola llamada
    Set authData = m_AuthRepository.GetUserAuthData(userEmail)

    If Not authData Is Nothing Then
        ' Verificar si el usuario existe
        If authData.UserExists Then
            ' Determinar el rol basándose en las propiedades de EAuthData
            ' Prioridad: Administrador Global > Administrador App > Calidad > Técnico
            If authData.IsGlobalAdmin Then
                role = RolAdmin
            ElseIf authData.IsAppAdmin Then
                role = RolAdmin
            ElseIf authData.IsCalidad Then
                role = RolCalidad
            ElseIf authData.IsTecnico Then
                role = RolTecnico
            Else
                role = RolDesconocido
            End If
        Else
            ' Usuario no encontrado
            role = RolDesconocido
        End If
    End If
    
    Set authData = Nothing

    ' Logging de auditoría: Resultado de consulta de rol
    If Not m_OperationLogger Is Nothing Then
        Dim roleText As String
        Select Case role
            Case RolAdmin: roleText = "Administrador"
            Case RolCalidad: roleText = "Calidad"
            Case RolTecnico: roleText = "Técnico"
            Case Else: roleText = "Desconocido"
        End Select
        m_OperationLogger.LogOperation "ROLE_ASSIGNED", 0, "Rol asignado a " & userEmail & ": " & roleText & " (CAuthService.GetUserRole)"
    End If
    
    IAuthService_GetUserRole = role
    Exit Function

errorHandler:
    IAuthService_GetUserRole = RolDesconocido
    If Not authData Is Nothing Then
        Set authData = Nothing
    End If
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_ERROR", 0, "Error consultando rol para " & userEmail & ": " & Err.Description & " (CAuthService.GetUserRole)"
    End If
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthService.GetUserRole"
    End If
End Function

Private Function IAuthService_AuthenticateUser(ByVal userEmail As String, ByVal password As String) As Boolean
    ' TODO: Implementar la lógica real de autenticación.
    ' Por ahora, devolvemos False para que compile.
    IAuthService_AuthenticateUser = False
End Function

' Método público para acceso directo
Public Function GetUserRole(ByVal userEmail As String) As UserRole
    GetUserRole = IAuthService_GetUserRole(userEmail)
End Function

Public Function AuthenticateUser(ByVal userEmail As String, ByVal password As String) As Boolean
    AuthenticateUser = IAuthService_AuthenticateUser(userEmail, password)
End Function

' Implementación del nuevo método de la interfaz
Private Function IAuthService_GetCurrentUserEmail() As String
    IAuthService_GetCurrentUserEmail = configSvc.GetUsuarioActual()
End Function

' Método público de conveniencia
Public Function GetCurrentUserEmail() As String
    GetCurrentUserEmail = IAuthService_GetCurrentUserEmail()
End Function















