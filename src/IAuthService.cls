Option Compare Database
Option Explicit

' Contrato para los servicios de autenticación y autorización.

Public Function AuthenticateUser(ByVal email As String) As Boolean
End Function

Public Function GetUserRole(ByVal email As String) As E_UserRole
End Function

Public Function IsUserAuthorized(ByVal email As String, ByVal permission As String) As Boolean
End Function

Public Function ValidateEmail(ByVal email As String) As Boolean
End Function





