Option Compare Database
Option Explicit
' Compilación condicional removida para evitar errores
' #If DEV_MODE Then
Implements IAuthService

' Implementación de la interfaz IAuthService
Public Sub Initialize(config As IConfig)
    ' Mock no requiere configuración específica, pero mantiene la interfaz
    ' para compatibilidad con el patrón de inyección de dependencias
End Sub

Private Function IAuthService_AuthenticateUser(ByVal Email As String) As Boolean
    Select Case LCase(Email)
        Case "admin.condor@example.com", "calidad.condor@example.com", "tecnico.condor@example.com"
            IAuthService_AuthenticateUser = True
        Case Else
            IAuthService_AuthenticateUser = False
    End Select
End Function

Private Function IAuthService_GetUserRole(ByVal Email As String) As E_UserRole
    Select Case LCase(Email)
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

Private Function IAuthService_IsUserAuthorized(ByVal Email As String, ByVal permission As String) As Boolean
    ' Lógica MOCK para autorización.
    ' Devuelve True solo si el permiso es "read" para cualquier usuario de prueba.
    If LCase(permission) = "read" Then
        IAuthService_IsUserAuthorized = True
    Else
        IAuthService_IsUserAuthorized = False
    End If
End Function

Private Function IAuthService_ValidateEmail(ByVal Email As String) As Boolean
    ' Lógica MOCK simple para validación
    IAuthService_ValidateEmail = (Email <> "email_invalido")
End Function

Private Function IAuthService_IsUserAuthenticated() As Boolean
    ' Lógica MOCK para verificar si un usuario está autenticado.
    ' Para propósitos de mock, asumimos que si se llamó a AuthenticateUser con éxito, está autenticado.
    IAuthService_IsUserAuthenticated = True ' Simplificado para el mock
End Function

Private Function IAuthService_GetAuthenticatedUserEmail() As String
    ' Lógica MOCK para obtener el email del usuario autenticado.
    ' Para propósitos de mock, devolvemos un email de prueba.
    IAuthService_GetAuthenticatedUserEmail = "mock.user@example.com"
End Function
' #End If












