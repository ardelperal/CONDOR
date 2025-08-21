Option Compare Database
Option Explicit
' Contrato para los servicios de autenticación y autorización.

Public Function AuthenticateUser(ByVal Email As String) As Boolean
End Function

Public Function GetUserRole(ByVal Email As String) As E_UserRole
End Function

Public Function IsUserAuthorized(ByVal Email As String, ByVal permission As String) As Boolean
End Function

Public Function ValidateEmail(ByVal Email As String) As Boolean
End Function

Public Function IsUserAuthenticated() As Boolean
End Function

Public Function GetAuthenticatedUserEmail() As String
End Function












