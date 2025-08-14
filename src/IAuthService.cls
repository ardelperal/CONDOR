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

' =====================================================
' INTERFAZ: IAuthService
' PROPOSITO: Define el contrato para servicios de autenticacion
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' =====================================================
' FUNCION: GetUserRole
' PROPOSITO: Determinar el rol de un usuario basado en su email
' PARAMETROS:
'   - userEmail: String - Email del usuario a verificar
' RETORNA: E_UserRole - Rol del usuario en el sistema
' =====================================================
Public Function GetUserRole(ByVal userEmail As String) As E_UserRole
    ' Esta es una interfaz - debe ser implementada por las clases concretas
End Function