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
    On Error GoTo ErrorHandler
    
    Set configSvc = config
    Set m_OperationLogger = operationLogger
    Set m_AuthRepository = authRepository
    Set m_ErrorHandler = errorHandler
    
    Exit Sub
    
ErrorHandler:
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthService.Initialize"
    End If
End Sub

Private Function IAuthService_GetUserRole(ByVal UserEmail As String) As E_UserRole
    On Error GoTo ErrorHandler

    ' Logging de auditoría: Consulta de rol de usuario
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_QUERY", "Consultando rol para usuario: " & UserEmail, "CAuthService.GetUserRole"
    End If

    Dim role As E_UserRole
    Dim authData As E_AuthData

    role = Rol_Desconocido

    ' Obtener todos los datos de autenticación en una sola llamada
    Set authData = m_AuthRepository.GetUserAuthData(UserEmail)

    If Not authData Is Nothing Then
        ' Verificar si el usuario existe
        If authData.UserExists Then
            ' Determinar el rol basándose en las propiedades de E_AuthData
            ' Prioridad: Administrador Global > Administrador App > Calidad > Técnico
            If authData.IsGlobalAdmin Then
                role = ROL_ADMINISTRADOR
            ElseIf authData.IsAppAdmin Then
                role = ROL_ADMINISTRADOR
            ElseIf authData.IsCalidad Then
                role = Rol_Calidad
            ElseIf authData.IsTecnico Then
                role = Rol_Tecnico
            Else
                role = Rol_Desconocido
            End If
        Else
            ' Usuario no encontrado
            role = Rol_Desconocido
        End If
    End If
    
    Set authData = Nothing

    ' Logging de auditoría: Resultado de consulta de rol
    If Not m_OperationLogger Is Nothing Then
        Dim roleText As String
        Select Case role
            Case ROL_ADMINISTRADOR: roleText = "Administrador"
            Case Rol_Calidad: roleText = "Calidad"
            Case Rol_Tecnico: roleText = "Técnico"
            Case Else: roleText = "Desconocido"
        End Select
        m_OperationLogger.LogOperation "ROLE_ASSIGNED", "Rol asignado a " & UserEmail & ": " & roleText, "CAuthService.GetUserRole"
    End If
    
    IAuthService_GetUserRole = role
    Exit Function

ErrorHandler:
    IAuthService_GetUserRole = Rol_Desconocido
    If Not authData Is Nothing Then
        Set authData = Nothing
    End If
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_ERROR", "Error consultando rol para " & UserEmail & ": " & Err.Description, "CAuthService.GetUserRole"
    End If
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthService.GetUserRole"
    End If
End Function

' Método público para acceso directo
Public Function GetUserRole(ByVal UserEmail As String) As E_UserRole
    GetUserRole = IAuthService_GetUserRole(UserEmail)
End Function













