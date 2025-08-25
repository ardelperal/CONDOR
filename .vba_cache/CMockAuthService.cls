Option Compare Database
Option Explicit

#If DEV_MODE Then

Implements IAuthService

' ImplementaciÃ³n de la interfaz IAuthService
Public Sub Initialize(config As IConfig)
    ' Mock no requiere configuraciÃ³n especÃ­fica, pero mantiene la interfaz
    ' para compatibilidad con el patrÃ³n de inyecciÃ³n de dependencias
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












