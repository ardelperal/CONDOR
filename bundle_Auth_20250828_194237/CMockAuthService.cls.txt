VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


#If DEV_MODE Then

Implements IAuthService

' Variable privada para almacenar el rol configurado
Private m_MockUserRole As E_UserRole

' Implementación de la interfaz IAuthService
Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, authRepository As IAuthRepository, errorHandler As IErrorHandlerService)
    ' Mock no requiere configuración específica, pero mantiene la interfaz
    ' para compatibilidad con el patrón de inyección de dependencias
    m_MockUserRole = Rol_Desconocido ' Valor por defecto
End Sub

' Método público para configurar el rol que debe devolver el mock
Public Sub SetMockUserRole(ByVal role As E_UserRole)
    m_MockUserRole = role
End Sub

Private Function IAuthService_GetUserRole(ByVal UserEmail As String) As E_UserRole
    ' El mock devuelve el rol configurado mediante SetMockUserRole
    ' independientemente del email proporcionado
    IAuthService_GetUserRole = m_MockUserRole
End Function

' Método público de conveniencia que delega a la implementación de interfaz
Public Function GetUserRole(ByVal UserEmail As String) As E_UserRole
    GetUserRole = IAuthService_GetUserRole(UserEmail)
End Function

' Método para resetear el estado del mock
Public Sub Reset()
    m_MockUserRole = Rol_Desconocido
End Sub

#End If














