VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthService"
Option Compare Database
Option Explicit

' =====================================================
' INTERFAZ: IAuthService
' PROPÓSITO: Define el contrato para servicios de autenticación
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' =====================================================
' FUNCIÓN: GetUserRole
' PROPÓSITO: Determinar el rol de un usuario basado en su email
' PARÁMETROS:
'   - userEmail: String - Email del usuario a verificar
' RETORNA: E_UserRole - Rol del usuario en el sistema
' =====================================================
Public Function GetUserRole(ByVal userEmail As String) As E_UserRole
    ' Esta es una interfaz - debe ser implementada por las clases concretas
End Function