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
Private m_SolicitudRepository As ISolicitudRepository

Private Sub Class_Initialize()
    ' Inicialización sin dependencias
End Sub

Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, solicitudRepository As ISolicitudRepository)
    On Error GoTo ErrorHandler
    
    Set configSvc = config
    Set m_OperationLogger = operationLogger
    Set m_SolicitudRepository = solicitudRepository
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CAuthService.Initialize"
End Sub

Private Function IAuthService_GetUserRole(ByVal UserEmail As String) As E_UserRole
    On Error GoTo ErrorHandler

    ' Logging de auditoría: Consulta de rol de usuario
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_QUERY", "Consultando rol para usuario: " & UserEmail, "CAuthService.GetUserRole"
    End If

    Dim role As E_UserRole
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim dbPath As String
    Dim dbPassword As String

    role = Rol_Desconocido

    dbPath = configSvc.GetValue("LANZADERADBPATH")
    dbPassword = configSvc.GetValue("DATABASEPASSWORD")

    If dbPath = "" Or dbPassword = "" Then
        Call modErrorHandler.LogWarning("La configuración para la base de datos de Lanzadera no fue encontrada.", "CAuthService.GetUserRole")
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If

    ' Verificar si es administrador global
    sql = "SELECT EsAdministrador FROM TbUsuariosAplicaciones WHERE CorreoUsuario = '" & UserEmail & "'"
    Set rs = ExecuteExternalQuery(sql, dbPath, dbPassword)

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            If rs!EsAdministrador = "Sí" Then
                role = Rol_Admin
            End If
        End If
        rs.Close
    End If

    If role = Rol_Admin Then
        IAuthService_GetUserRole = Rol_Admin
        Exit Function
    End If

    ' Consultar roles específicos de la aplicación
    sql = "SELECT EsUsuarioAdministrador, EsUsuarioCalidad, EsUsuarioTecnico " & _
          "FROM TbUsuariosAplicacionesPermisos " & _
          "WHERE CorreoUsuario = '" & UserEmail & "' AND IDAplicacion = " & modConfig.IDAplicacion_CONDOR
    Set rs = ExecuteExternalQuery(sql, dbPath, dbPassword)

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            If rs!EsUsuarioAdministrador = "Sí" Then
                role = Rol_Admin
            ElseIf rs!EsUsuarioCalidad = "Sí" Then
                role = Rol_Calidad
            ElseIf rs!EsUsuarioTecnico = "Sí" Then
                role = Rol_Tecnico
            End If
        End If
        rs.Close
    End If

    ' Logging de auditoría: Resultado de consulta de rol
    If Not m_OperationLogger Is Nothing Then
        Dim roleText As String
        Select Case role
            Case Rol_Admin: roleText = "Administrador"
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
    If Not rs Is Nothing Then Set rs = Nothing
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "ROLE_ERROR", "Error consultando rol para " & UserEmail & ": " & Err.Description, "CAuthService.GetUserRole"
    End If
    Call modErrorHandler.LogError(Err.Number, Err.Description, "CAuthService.GetUserRole")
End Function

' Método público para acceso directo
Public Function GetUserRole(ByVal UserEmail As String) As E_UserRole
    GetUserRole = IAuthService_GetUserRole(UserEmail)
End Function











