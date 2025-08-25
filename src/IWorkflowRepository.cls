VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Interfaz: IWorkflowRepository
' Propósito: Define el contrato para gestionar las reglas de transición de estados
'            del workflow de solicitudes
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
'==============================================================================

Option Explicit

'==============================================================================
' MÉTODOS PÚBLICOS
'==============================================================================

' Verifica si una transición de estado es válida según las reglas del workflow
' @param tipoSolicitud: Tipo de solicitud (ej: "PC", "CD_CA")
' @param estadoOrigen: Estado actual de la solicitud
' @param estadoDestino: Estado al que se quiere transicionar
' @return: True si la transición es válida, False en caso contrario
Public Function IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
End Function
