VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =====================================================
' Interfaz: IAuthRepository
' Propósito: Define el contrato para el acceso a datos de autenticación
' Autor: CONDOR-Expert
' Fecha: 2025-01-15
' =====================================================


' Obtiene todos los datos de autenticación de un usuario en una sola consulta
' @param params: Collection de objetos QueryParameter con los parámetros de la consulta
' @return: Objeto E_AuthData con toda la información de autenticación
Public Function GetUserAuthData(ByVal params As Collection) As E_AuthData
End Function

