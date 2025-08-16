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

Public Function AuthenticateUser(email As String) As Boolean
End Function

Public Function IsUserAuthorized(email As String, permission As String) As Boolean
End Function

Public Function GetUserRole(email As String) As E_UserRole
End Function





