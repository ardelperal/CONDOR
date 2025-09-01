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

' Miembros privados para el mock de GetUserRole
Private m_getUserRoleReturnValue As UserRole
Private m_getUserRoleWasCalled As Boolean
Private m_lastUserEmailForRole As String

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' Implementación de la interfaz IAuthService
Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, authRepository As IAuthRepository, errorHandler As IErrorHandlerService)
    ' Mock no necesita inicialización real
End Sub

' Configuración del valor de retorno para GetUserRole
Public Sub ConfigureGetUserRole(ByVal returnValue As UserRole)
    m_getUserRoleReturnValue = returnValue
End Sub

' Propiedades "espía" para GetUserRole
Public Property Get GetUserRole_WasCalled() As Boolean
    GetUserRole_WasCalled = m_getUserRoleWasCalled
End Property

Public Property Get LastUserEmailForRole() As String
    LastUserEmailForRole = m_lastUserEmailForRole
End Property

Private Function IAuthService_GetUserRole(ByVal userEmail As String) As UserRole
    m_getUserRoleWasCalled = True
    m_lastUserEmailForRole = userEmail
    IAuthService_GetUserRole = m_getUserRoleReturnValue
End Function

' Método público de conveniencia que delega a la implementación de interfaz
Public Function GetUserRole(ByVal userEmail As String) As UserRole
    GetUserRole = IAuthService_GetUserRole(userEmail)
End Function

' Método para resetear el estado del mock
Public Sub Reset()
    m_getUserRoleReturnValue = RolDesconocido
    m_getUserRoleWasCalled = False
    m_lastUserEmailForRole = ""
End Sub