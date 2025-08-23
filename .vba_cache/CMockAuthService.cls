Option Compare Database
Option Explicit

#If DEV_MODE Then

Implements IAuthService

' Implementación de la interfaz IAuthService
Public Sub Initialize(config As IConfig)
    ' Mock no requiere configuración específica, pero mantiene la interfaz
    ' para compatibilidad con el patrón de inyección de dependencias
End Sub

Private Function IAuthService_GetUserRole(ByVal UserEmail As String) As E_UserRole
    Select Case LCase(UserEmail)
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

#End If












