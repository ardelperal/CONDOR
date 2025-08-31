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

Implements IAuthService

' Variable privada para almacenar el rol configurado
Private m_MockUserRole As UserRole

' Implementación de la interfaz IAuthService
Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, authRepository As IAuthRepository, errorHandler As IErrorHandlerService)
    m_MockUserRole = RolDesconocido ' Valor por defecto
End Sub

' Método público para configurar el rol que debe devolver el mock
Public Sub SetMockUserRole(ByVal role As UserRole)
    m_MockUserRole = role
End Sub

Private Function IAuthService_GetUserRole(ByVal userEmail As String) As UserRole
    IAuthService_GetUserRole = m_MockUserRole
End Function

' Método público de conveniencia que delega a la implementación de interfaz
Public Function GetUserRole(ByVal userEmail As String) As UserRole
    GetUserRole = IAuthService_GetUserRole(userEmail)
End Function

' Método para resetear el estado del mock
Public Sub Reset()
    m_MockUserRole = RolDesconocido
End Sub