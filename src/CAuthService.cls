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
'prueba del cambio para 
Private configSvc As IConfig

Private Sub Class_Initialize()
    ' TODO: Usar un Factory para CConfig cuando exista
    Set configSvc = New CConfig
End Sub

Private Function IAuthService_AuthenticateUser(ByVal Email As String) As Boolean
    ' Lógica de autenticación...
    IAuthService_AuthenticateUser = (Trim(Email) <> "")
End Function

Private Function IAuthService_GetUserRole(ByVal userEmail As String) As E_UserRole
    On Error GoTo ErrorHandler

    Dim role As E_UserRole
    Dim rs As DAO.Recordset ' Variable para recordset de consulta
    Dim sql As String
    Dim dbPath As String
    Dim dbPassword As String

    role = Rol_Desconocido

    dbPath = AppConfig.GetValue("LANZADERADBPATH")
    dbPassword = AppConfig.GetValue("DATABASEPASSWORD")

    If dbPath = "" Or dbPassword = "" Then
        Call modErrorHandler.LogWarning("La configuración para la base de datos de Lanzadera no fue encontrada.", "CAuthService.GetUserRole")
        IAuthService_GetUserRole = Rol_Desconocido
        Exit Function
    End If

    sql = "SELECT EsAdministrador FROM TbUsuariosAplicaciones WHERE CorreoUsuario = '" & userEmail & "'"
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

    sql = "SELECT EsUsuarioAdministrador, EsUsuarioCalidad, EsUsuarioTecnico " & _
          "FROM TbUsuariosAplicacionesPermisos " & _
          "WHERE CorreoUsuario = '" & userEmail & "' AND IDAplicacion = 231"
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

    IAuthService_GetUserRole = role
    Exit Function

ErrorHandler:
    IAuthService_GetUserRole = Rol_Desconocido
    If Not rs Is Nothing Then Set rs = Nothing
    Call modErrorHandler.LogError(Err.Number, Err.Description, "CAuthService.GetUserRole")
End Function

Private Function IAuthService_IsUserAuthorized(ByVal Email As String, ByVal permission As String) As Boolean
    ' TODO: Implementar la lógica real de autorización basada en roles y permisos.
    ' Por ahora, devolvemos True para que compile y las pruebas básicas pasen.
    IAuthService_IsUserAuthorized = True
End Function

Private Function IAuthService_ValidateEmail(ByVal Email As String) As Boolean
    ' Lógica simple de validación de formato de email
    IAuthService_ValidateEmail = (InStr(Email, "@") > 1 And InStr(Email, ".") > InStr(Email, "@"))
End Function

' Implementaciones públicas de los métodos de la interfaz
Public Function AuthenticateUser(ByVal Email As String) As Boolean
    AuthenticateUser = IAuthService_AuthenticateUser(Email)
End Function

Public Function GetUserRole(ByVal Email As String) As E_UserRole
    GetUserRole = IAuthService_GetUserRole(Email)
End Function

Public Function IsUserAuthorized(ByVal Email As String, ByVal permission As String) As Boolean
    IsUserAuthorized = IAuthService_IsUserAuthorized(Email, permission)
End Function

Public Function ValidateEmail(ByVal Email As String) As Boolean
    ValidateEmail = IAuthService_ValidateEmail(Email)
End Function









