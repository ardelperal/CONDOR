Option Compare Database
Option Explicit


' =====================================================
' INTERFAZ: IExpedienteRepository
' DESCRIPCIÓN: Define el contrato para el repositorio de expedientes.
'              Proporciona métodos para obtener expedientes.
' AUTOR: CONDOR-Architect
' FECHA: 2025-08-28
' =====================================================

Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
End Function

Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
End Function

Public Function ObtenerExpedientesActivosParaSelector() As Object
End Function