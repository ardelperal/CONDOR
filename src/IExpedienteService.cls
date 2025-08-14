VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' Interfaz: IExpedienteService
' Descripción: Define el contrato para servicios de expedientes
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Obtiene un expediente por su ID desde la aplicación de Expedientes existente
' @param idExpediente: ID del expediente a buscar
' @return: Estructura T_Expediente con los datos del expediente
Public Function GetExpedientePorID(ByVal idExpediente As Long) As T_Expediente
    ' Esta función debe ser implementada por las clases concretas
End Function

#End If