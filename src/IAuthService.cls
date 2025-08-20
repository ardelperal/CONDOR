VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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











