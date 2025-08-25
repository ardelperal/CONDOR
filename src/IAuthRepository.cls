VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================
' Interfaz: IAuthRepository
' Propósito: Define el contrato para el acceso a datos de autenticación
' Autor: CONDOR-Expert
' Fecha: 2025-01-15
' =====================================================

Option Explicit

' Obtiene los datos de un usuario por su email
' @param UserEmail: Email del usuario a buscar
' @return: Recordset con los datos del usuario o Nothing si no existe
Public Function GetUserByEmail(ByVal UserEmail As String) As DAO.Recordset
End Function

' Obtiene los permisos específicos de aplicación para un usuario
' @param UserEmail: Email del usuario
' @param ApplicationID: ID de la aplicación
' @return: Recordset con los permisos del usuario o Nothing si no existe
Public Function GetUserApplicationPermissions(ByVal UserEmail As String, ByVal ApplicationID As String) As DAO.Recordset
End Function