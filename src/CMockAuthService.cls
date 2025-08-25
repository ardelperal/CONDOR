Option Compare Database
Option Explicit

#If DEV_MODE Then

Implements IAuthService

' Variable privada para almacenar el rol configurado
Private m_MockUserRole As E_UserRole

' Implementación de la interfaz IAuthService
Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, authRepository As IAuthRepository)
    ' Mock no requiere configuración específica, pero mantiene la interfaz
    ' para compatibilidad con el patrón de inyección de dependencias
    m_MockUserRole = Rol_Desconocido ' Valor por defecto
End Sub

' MÃ©todo pÃºblico para configurar el rol que debe devolver el mock
Public Sub SetMockUserRole(ByVal role As E_UserRole)
    m_MockUserRole = role
End Sub

Private Function IAuthService_GetUserRole(ByVal UserEmail As String) As E_UserRole
    ' El mock devuelve el rol configurado mediante SetMockUserRole
    ' independientemente del email proporcionado
    IAuthService_GetUserRole = m_MockUserRole
End Function

' MÃ©todo pÃºblico de conveniencia que delega a la implementaciÃ³n de interfaz
Public Function GetUserRole(ByVal UserEmail As String) As E_UserRole
    GetUserRole = IAuthService_GetUserRole(UserEmail)
End Function

#End If












