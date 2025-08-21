Option Compare Database
Option Explicit
' Compilación condicional removida para evitar errores
' #If DEV_MODE Then

' ============================================================================
' Interfaz: IExpedienteService
' Descripci?n: Define el contrato para servicios de expedientes
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Obtiene un expediente por su ID desde la aplicación de Expedientes existente
' @param idExpediente: ID del expediente a buscar
' @return: Estructura T_Expediente con los datos del expediente
Public Function GetExpedienteById(ByVal idExpediente As Long) As T_Expediente
    ' Esta función debe ser implementada por las clases concretas
End Function

' Métodos adicionales identificados en las pruebas
Public Function GetExpediente(ByVal ID As Long) As T_Expediente
End Function

Public Function CreateExpediente(ByVal NumeroExpediente As String, ByVal Descripcion As String, ByVal idUsuario As Long) As Long
End Function

Public Function UpdateExpediente(ByVal ID As Long, ByVal Descripcion As String, ByVal Estado As String) As Boolean
End Function

Public Function DeleteExpediente(ByVal ID As Long) As Boolean
End Function

Public Function SearchExpedientes(ByVal criterio As String) As Collection
End Function

Public Function GetExpedientesByUser(ByVal idUsuario As Long) As Collection
End Function

Public Function ValidateExpediente(expediente As T_Expediente) As Boolean
End Function

' #End If











