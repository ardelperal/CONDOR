VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExpedienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' INTERFAZ: IExpedienteRepository
' Arquitectura: Contrato Limpio para Acceso a Datos de Expedientes
' Versión: 1.0 - Creación Inicial
' ============================================================================
' Define el contrato para operaciones de persistencia de expedientes.
' Implementa el patrón Repository con conexión explícita al Backend.
' Cumple con el Principio de Responsabilidad Única: solo gestiona datos de expedientes.
' ============================================================================

' Obtiene un expediente por su ID desde el Backend
' @param IdExpediente: ID único del expediente
' @return: Recordset con los datos del expediente o Nothing si no existe
Public Function ObtenerExpedientePorId(ByVal IdExpediente As Long) As DAO.Recordset
End Function

' Obtiene un expediente por su nemotécnico desde el Backend
' @param Nemotecnico: Nemotécnico único del expediente
' @return: Recordset con los datos del expediente o Nothing si no existe
Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As DAO.Recordset
End Function

' Obtiene la lista de expedientes activos para selectores
' @return: Recordset con expedientes en estado 'Activo' o 'En Proceso'
Public Function ObtenerExpedientesActivosParaSelector() As DAO.Recordset
End Function