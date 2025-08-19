#If DEV_MODE Then
Option Compare Database
Option Explicit
Implements IAuthService

Private Function IAuthService_AuthenticateUser(ByVal email As String) As Boolean
    Select Case LCase(email)
        Case "admin.condor@example.com", "calidad.condor@example.com", "tecnico.condor@example.com"
            IAuthService_AuthenticateUser = True
        Case Else
            IAuthService_AuthenticateUser = False
    End Select
End Function

Private Function IAuthService_GetUserRole(ByVal email As String) As E_UserRole
    Select Case LCase(email)
        Case "admin.condor@example.com"
            IAuthService_GetUserRole = Rol_Admin
        Case "calidad.condor@example.com"
            IAuthService_GetUserRole = Rol_Calidad
        Case "tecnico.condor@example.com"
            IAuthService_GetUserRole = Rol_Tecnico
        Case Else
            IAuthService_GetUserRole = Rol_Desconocido
    End Select
End Function

Private Function IAuthService_IsUserAuthorized(ByVal email As String, ByVal permission As String) As Boolean
    ' L├│gica MOCK para autorizaci├│n.
    ' Devuelve True solo si el permiso es "read" para cualquier usuario de prueba.
    If LCase(permission) = "read" Then
        IAuthService_IsUserAuthorized = True
    Else
        IAuthService_IsUserAuthorized = False
    End If
End Function

Private Function IAuthService_ValidateEmail(ByVal email As String) As Boolean
    ' L├│gica MOCK simple para validaci├│n
    IAuthService_ValidateEmail = (email <> "email_invalido")
End Function
#End If





