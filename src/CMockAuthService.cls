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
#End If











